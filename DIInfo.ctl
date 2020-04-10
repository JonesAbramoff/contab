VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl DIInfo 
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9720
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   9720
   Begin VB.Frame FrameOpcao 
      BorderStyle     =   0  'None
      Height          =   6045
      Index           =   1
      Left            =   135
      TabIndex        =   63
      Top             =   600
      Width           =   9300
      Begin VB.Frame Frame12 
         Caption         =   "Desembaraço Aduaneiro"
         Height          =   585
         Left            =   15
         TabIndex        =   180
         Top             =   1815
         Width           =   9300
         Begin VB.ComboBox DUF 
            Height          =   315
            Left            =   3420
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   195
            Width           =   855
         End
         Begin VB.TextBox DLocal 
            Height          =   315
            Left            =   4980
            MaxLength       =   60
            TabIndex        =   12
            Top             =   180
            Width           =   4215
         End
         Begin MSMask.MaskEdBox DData 
            Height          =   315
            Left            =   1470
            TabIndex        =   9
            Top             =   210
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDData 
            Height          =   300
            Left            =   2685
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   210
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
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
            Height          =   180
            Left            =   885
            TabIndex        =   183
            Top             =   255
            Width           =   525
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   3045
            TabIndex        =   182
            Top             =   255
            Width           =   315
         End
         Begin VB.Label Label41 
            Caption         =   "Local:"
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
            Height          =   240
            Left            =   4380
            TabIndex        =   181
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Incluir dados do Pedido de Compra"
         Height          =   915
         Left            =   30
         TabIndex        =   170
         Top             =   5085
         Width           =   9285
         Begin VB.CommandButton BotaoIncluirPC 
            Caption         =   "Incluir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   8025
            TabIndex        =   35
            Top             =   195
            Width           =   690
         End
         Begin MSMask.MaskEdBox PC 
            Height          =   300
            Left            =   1440
            TabIndex        =   34
            Top             =   225
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin VB.Label PCFilial 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4350
            TabIndex        =   179
            Top             =   555
            Width           =   2025
         End
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
            Left            =   3780
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   178
            Top             =   615
            Width           =   465
         End
         Begin VB.Label PCValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6825
            TabIndex        =   177
            Top             =   210
            Width           =   1155
         End
         Begin VB.Label Label42 
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
            Left            =   6255
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   176
            Top             =   270
            Width           =   510
         End
         Begin VB.Label PCData 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4350
            TabIndex        =   175
            Top             =   210
            Width           =   1155
         End
         Begin VB.Label Label40 
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
            Left            =   3495
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   174
            Top             =   270
            Width           =   750
         End
         Begin VB.Label PCForn 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1440
            TabIndex        =   173
            Top             =   555
            Width           =   2205
         End
         Begin VB.Label Label33 
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
            Left            =   285
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   172
            Top             =   600
            Width           =   1035
         End
         Begin VB.Label PCLabel 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   390
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   171
            Top             =   270
            Width           =   930
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Valores\Peso"
         Height          =   2175
         Left            =   30
         TabIndex        =   117
         Top             =   2910
         Width           =   9285
         Begin VB.ComboBox MoedaItens 
            Height          =   315
            ItemData        =   "DIInfo.ctx":0000
            Left            =   8130
            List            =   "DIInfo.ctx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   180
            Width           =   930
         End
         Begin VB.ComboBox Moeda2 
            Height          =   315
            ItemData        =   "DIInfo.ctx":0014
            Left            =   1440
            List            =   "DIInfo.ctx":0016
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   540
            Width           =   2400
         End
         Begin VB.CommandButton BotaoTrazCotacao2 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6510
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Numeração Automática"
            Top             =   525
            Width           =   345
         End
         Begin VB.Frame Frame8 
            Caption         =   "Pesos KG"
            Height          =   1275
            Left            =   7260
            TabIndex        =   128
            Top             =   840
            Width           =   1920
            Begin MSMask.MaskEdBox DIPesoBruto 
               Height          =   315
               Left            =   855
               TabIndex        =   32
               Top             =   315
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox DIPesoLiquido 
               Height          =   315
               Left            =   855
               TabIndex        =   33
               Top             =   765
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label17 
               Caption         =   "Bruto:"
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
               Left            =   300
               TabIndex        =   130
               Top             =   375
               Width           =   690
            End
            Begin VB.Label Label25 
               Caption         =   "Líquido:"
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
               TabIndex        =   129
               Top             =   825
               Width           =   885
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Valores em R$"
            Height          =   1275
            Left            =   4605
            TabIndex        =   122
            Top             =   840
            Width           =   2490
            Begin MSMask.MaskEdBox ValorMercadoriaEmReal 
               Height          =   315
               Left            =   1200
               TabIndex        =   29
               Top             =   210
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ValorFreteInternacEmReal 
               Height          =   315
               Left            =   1200
               TabIndex        =   30
               Top             =   555
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ValorSeguroInternacEmReal 
               Height          =   315
               Left            =   1200
               TabIndex        =   31
               Top             =   900
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label8 
               Caption         =   "Mercadoria:"
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
               Left            =   75
               TabIndex        =   125
               Top             =   285
               Width           =   1005
            End
            Begin VB.Label Label7 
               Caption         =   "Frete:"
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
               Left            =   585
               TabIndex        =   124
               Top             =   570
               Width           =   540
            End
            Begin VB.Label Label6 
               Caption         =   "Seguro:"
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
               Left            =   420
               TabIndex        =   123
               Top             =   915
               Width           =   705
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Valores em Moeda"
            Height          =   1275
            Left            =   285
            TabIndex        =   118
            Top             =   840
            Width           =   4140
            Begin VB.ComboBox MoedaSeguro 
               Height          =   315
               ItemData        =   "DIInfo.ctx":0018
               Left            =   3120
               List            =   "DIInfo.ctx":0022
               Style           =   2  'Dropdown List
               TabIndex        =   28
               Top             =   900
               Width           =   930
            End
            Begin VB.ComboBox MoedaFrete 
               Height          =   315
               ItemData        =   "DIInfo.ctx":002C
               Left            =   3120
               List            =   "DIInfo.ctx":0036
               Style           =   2  'Dropdown List
               TabIndex        =   26
               Top             =   540
               Width           =   930
            End
            Begin VB.ComboBox MoedaMercadoria 
               Height          =   315
               ItemData        =   "DIInfo.ctx":0040
               Left            =   3120
               List            =   "DIInfo.ctx":004A
               Style           =   2  'Dropdown List
               TabIndex        =   24
               Top             =   195
               Width           =   930
            End
            Begin MSMask.MaskEdBox ValorMercadoriaMoeda 
               Height          =   315
               Left            =   1155
               TabIndex        =   23
               Top             =   195
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ValorFreteInternacMoeda 
               Height          =   315
               Left            =   1155
               TabIndex        =   25
               Top             =   540
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ValorSeguroInternacMoeda 
               Height          =   315
               Left            =   1155
               TabIndex        =   27
               Top             =   900
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label50 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2430
               TabIndex        =   189
               Top             =   945
               Width           =   645
            End
            Begin VB.Label Label49 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2430
               TabIndex        =   188
               Top             =   585
               Width           =   645
            End
            Begin VB.Label Label48 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2430
               TabIndex        =   187
               Top             =   240
               Width           =   645
            End
            Begin VB.Label Label5 
               Caption         =   "Seguro:"
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
               Left            =   375
               TabIndex        =   121
               Top             =   915
               Width           =   705
            End
            Begin VB.Label Label4 
               Caption         =   "Frete:"
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
               Left            =   540
               TabIndex        =   120
               Top             =   555
               Width           =   540
            End
            Begin VB.Label Label3 
               Caption         =   "Mercadoria:"
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
               Left            =   45
               TabIndex        =   119
               Top             =   240
               Width           =   1005
            End
         End
         Begin VB.CommandButton BotaoTrazCotacao1 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6510
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Numeração Automática"
            Top             =   180
            Width           =   345
         End
         Begin VB.ComboBox Moeda1 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   195
            Width           =   2400
         End
         Begin MSMask.MaskEdBox TaxaMoeda1 
            Height          =   315
            Left            =   4965
            TabIndex        =   17
            Top             =   195
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            Format          =   "###,##0.00####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TaxaMoeda2 
            Height          =   315
            Left            =   4965
            TabIndex        =   20
            Top             =   540
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            Format          =   "###,##0.00####"
            PromptChar      =   " "
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "Moeda Itens:"
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
            Left            =   6960
            TabIndex        =   186
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "Taxa 2:"
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
            Left            =   4245
            TabIndex        =   185
            Top             =   615
            Width           =   660
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Moeda 2:"
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
            Left            =   555
            TabIndex        =   184
            Top             =   600
            Width           =   810
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Moeda 1:"
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
            Left            =   555
            TabIndex        =   127
            Top             =   255
            Width           =   810
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Taxa 1:"
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
            Left            =   4245
            TabIndex        =   126
            Top             =   270
            Width           =   660
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Identificação"
         Height          =   1770
         Left            =   15
         TabIndex        =   111
         Top             =   30
         Width           =   9300
         Begin VB.ComboBox ViaTransp 
            Height          =   315
            ItemData        =   "DIInfo.ctx":0054
            Left            =   6585
            List            =   "DIInfo.ctx":0077
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   855
            Width           =   2625
         End
         Begin VB.Frame FrameAdquirente 
            Caption         =   "Adquirente ou Encomendante"
            Height          =   570
            Left            =   4935
            TabIndex        =   213
            Top             =   1125
            Width           =   4260
            Begin VB.ComboBox UFAdquir 
               Height          =   315
               Left            =   3315
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   180
               Width           =   855
            End
            Begin MSMask.MaskEdBox CNPJAdquir 
               Height          =   315
               Left            =   660
               TabIndex        =   7
               Top             =   195
               Width           =   2040
               _ExtentX        =   3598
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   14
               Mask            =   "##############"
               PromptChar      =   " "
            End
            Begin VB.Label Label57 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2940
               TabIndex        =   215
               Top             =   240
               Width           =   315
            End
            Begin VB.Label Label56 
               AutoSize        =   -1  'True
               Caption         =   "CNPJ:"
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
               TabIndex        =   214
               Top             =   240
               Width           =   540
            End
         End
         Begin VB.ComboBox Intermedio 
            Height          =   315
            ItemData        =   "DIInfo.ctx":0118
            Left            =   1485
            List            =   "DIInfo.ctx":0125
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1200
            Width           =   3360
         End
         Begin VB.TextBox CodExportador 
            Height          =   315
            Left            =   1485
            MaxLength       =   60
            TabIndex        =   4
            Top             =   855
            Width           =   3360
         End
         Begin VB.TextBox DIDescricao 
            Height          =   315
            Left            =   1485
            TabIndex        =   3
            Top             =   510
            Width           =   7710
         End
         Begin MSMask.MaskEdBox Numero 
            Height          =   315
            Left            =   1485
            TabIndex        =   0
            Top             =   165
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   12
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   315
            Left            =   3405
            TabIndex        =   1
            Top             =   165
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownData 
            Height          =   300
            Left            =   4620
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   165
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "Via de Transporte:"
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
            Left            =   4950
            TabIndex        =   211
            Top             =   900
            Width           =   1590
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "Intermédio:"
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
            Left            =   465
            TabIndex        =   212
            Top             =   1245
            Width           =   945
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            Caption         =   "Cód.Exportador:"
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
            Height          =   210
            Left            =   45
            TabIndex        =   193
            Top             =   900
            Width           =   1380
         End
         Begin VB.Label DIStatus 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7515
            TabIndex        =   116
            Top             =   150
            Width           =   1680
         End
         Begin VB.Label LabelNumero 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   330
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   115
            Top             =   210
            Width           =   1095
         End
         Begin VB.Label LabelData 
            Alignment       =   1  'Right Justify
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
            Height          =   180
            Left            =   2805
            TabIndex        =   114
            Top             =   210
            Width           =   525
         End
         Begin VB.Label LabelStatus 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Left            =   6645
            TabIndex        =   113
            Top             =   195
            Width           =   810
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
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
            Height          =   210
            Left            =   405
            TabIndex        =   112
            Top             =   555
            Width           =   1020
         End
      End
      Begin VB.Frame Frame1 
         Height          =   555
         Index           =   0
         Left            =   15
         TabIndex        =   64
         Top             =   2355
         Width           =   9300
         Begin VB.TextBox ProcessoTrading 
            Height          =   315
            Left            =   7500
            TabIndex        =   15
            Top             =   165
            Width           =   1695
         End
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   4980
            TabIndex        =   14
            Top             =   165
            Width           =   1575
         End
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   315
            Left            =   1470
            TabIndex        =   13
            Top             =   165
            Width           =   2790
            _ExtentX        =   4921
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label2 
            Caption         =   "Processo:"
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
            Left            =   6600
            TabIndex        =   67
            Top             =   210
            Width           =   885
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
            Left            =   345
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   66
            Top             =   195
            Width           =   1035
         End
         Begin VB.Label Label1 
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
            Index           =   15
            Left            =   4440
            TabIndex        =   65
            Top             =   210
            Width           =   465
         End
      End
   End
   Begin VB.Frame FrameOpcao 
      BorderStyle     =   0  'None
      Caption         =   "Frame13"
      Height          =   5940
      Index           =   6
      Left            =   135
      TabIndex        =   200
      Top             =   660
      Visible         =   0   'False
      Width           =   9360
      Begin VB.CommandButton BotaoLimparGridPC 
         Caption         =   "Limpar"
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
         Left            =   5745
         TabIndex        =   210
         Top             =   5535
         Width           =   960
      End
      Begin VB.CommandButton BotaoItensPC 
         Caption         =   "Pedido de Compra"
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
         Left            =   6960
         TabIndex        =   208
         Top             =   5520
         Width           =   2235
      End
      Begin VB.Frame Frame13 
         Caption         =   "Itens de Pedidos de Compra"
         Height          =   5460
         Left            =   120
         TabIndex        =   201
         Top             =   0
         Width           =   9135
         Begin MSMask.MaskEdBox DataPC 
            Height          =   225
            Left            =   1200
            TabIndex        =   209
            Top             =   1560
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.TextBox DescProdutoPC 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   3495
            MaxLength       =   50
            TabIndex        =   203
            Top             =   1035
            Width           =   2805
         End
         Begin MSMask.MaskEdBox UMPC 
            Height          =   225
            Left            =   6720
            TabIndex        =   204
            Top             =   1530
            Width           =   645
            _ExtentX        =   1138
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
         Begin MSMask.MaskEdBox QuantPC 
            Height          =   225
            Left            =   7320
            TabIndex        =   205
            Top             =   1080
            Width           =   990
            _ExtentX        =   1746
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
         Begin MSMask.MaskEdBox ProdutoPC 
            Height          =   225
            Left            =   2190
            TabIndex        =   206
            Top             =   990
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodigoPC 
            Height          =   225
            Left            =   615
            TabIndex        =   207
            Top             =   1050
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridItensPC 
            Height          =   4005
            Left            =   120
            TabIndex        =   202
            Top             =   240
            Width           =   8820
            _ExtentX        =   15558
            _ExtentY        =   7064
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame FrameOpcao 
      BorderStyle     =   0  'None
      Height          =   6150
      Index           =   3
      Left            =   135
      TabIndex        =   49
      Top             =   600
      Visible         =   0   'False
      Width           =   9360
      Begin VB.Frame Frame10 
         Caption         =   "Total da DI"
         Height          =   1170
         Left            =   4170
         TabIndex        =   155
         Top             =   4935
         Width           =   4020
         Begin VB.Label LabelPesoLiqDI 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2700
            TabIndex        =   197
            Top             =   840
            Width           =   1170
         End
         Begin VB.Label Label52 
            Caption         =   "Peso Liq.:"
            Height          =   240
            Left            =   1950
            TabIndex        =   196
            Top             =   885
            Width           =   705
         End
         Begin VB.Label Label39 
            Caption         =   "FOB:"
            Height          =   210
            Left            =   315
            TabIndex        =   163
            Top             =   270
            Width           =   405
         End
         Begin VB.Label Label38 
            Caption         =   "FOB R$:"
            Height          =   240
            Left            =   75
            TabIndex        =   162
            Top             =   600
            Width           =   675
         End
         Begin VB.Label LabelTotalFOBDIMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   735
            TabIndex        =   161
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label LabelTotalFOBDIReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   735
            TabIndex        =   160
            Top             =   555
            Width           =   1230
         End
         Begin VB.Label Label35 
            Caption         =   "CIF:"
            Height          =   210
            Left            =   2340
            TabIndex        =   159
            Top             =   240
            Width           =   315
         End
         Begin VB.Label Label34 
            Caption         =   "CIF R$:"
            Height          =   240
            Left            =   2085
            TabIndex        =   158
            Top             =   570
            Width           =   570
         End
         Begin VB.Label LabelTotalCIFDIMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2700
            TabIndex        =   157
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label LabelTotalCIFDIReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2700
            TabIndex        =   156
            Top             =   525
            Width           =   1170
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Total da Adição"
         Height          =   1185
         Left            =   75
         TabIndex        =   146
         Top             =   4920
         Width           =   4020
         Begin VB.Label LabelPesoLiqAdicao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2685
            TabIndex        =   199
            Top             =   840
            Width           =   1170
         End
         Begin VB.Label Label53 
            Caption         =   "Peso Liq.:"
            Height          =   240
            Left            =   1935
            TabIndex        =   198
            Top             =   885
            Width           =   705
         End
         Begin VB.Label LabelTotalCIFAdicaoReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2700
            TabIndex        =   154
            Top             =   525
            Width           =   1170
         End
         Begin VB.Label LabelTotalCIFAdicaoMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2700
            TabIndex        =   153
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label Label32 
            Caption         =   "CIF R$:"
            Height          =   240
            Left            =   2085
            TabIndex        =   152
            Top             =   570
            Width           =   570
         End
         Begin VB.Label Label31 
            Caption         =   "CIF:"
            Height          =   210
            Left            =   2340
            TabIndex        =   151
            Top             =   240
            Width           =   315
         End
         Begin VB.Label LabelTotalFOBAdicaoReal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   735
            TabIndex        =   150
            Top             =   555
            Width           =   1230
         End
         Begin VB.Label LabelTotalFOBAdicaoMoeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   735
            TabIndex        =   149
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label Label29 
            Caption         =   "FOB R$:"
            Height          =   240
            Left            =   75
            TabIndex        =   148
            Top             =   600
            Width           =   675
         End
         Begin VB.Label Label28 
            Caption         =   "FOB:"
            Height          =   210
            Left            =   315
            TabIndex        =   147
            Top             =   270
            Width           =   405
         End
      End
      Begin VB.CommandButton BotaoProduto 
         Caption         =   "Produto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   8265
         TabIndex        =   105
         Top             =   5040
         Width           =   1005
      End
      Begin VB.Frame FrameGridItens 
         Caption         =   "Itens das Adições"
         Height          =   4875
         Left            =   75
         TabIndex        =   50
         Top             =   30
         Width           =   9060
         Begin MSMask.MaskEdBox IPIValorUnitario 
            Height          =   300
            Left            =   5550
            TabIndex        =   195
            Top             =   1200
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin VB.CheckBox TotalCIFEmRealManual 
            Height          =   210
            Left            =   2835
            TabIndex        =   131
            Top             =   450
            Width           =   810
         End
         Begin VB.ComboBox AdicaoItem 
            Height          =   315
            Left            =   495
            Style           =   2  'Dropdown List
            TabIndex        =   110
            Top             =   795
            Width           =   690
         End
         Begin VB.ComboBox UnidadeMed 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4455
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox Descricao 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   2790
            TabIndex        =   68
            Top             =   135
            Width           =   1950
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   300
            Left            =   1185
            TabIndex        =   51
            Top             =   165
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   5880
            TabIndex        =   52
            Top             =   690
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorUnitFOBNaMoeda 
            Height          =   300
            Left            =   4770
            TabIndex        =   53
            Top             =   1095
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorUnitFOBEmReal 
            Height          =   300
            Left            =   4200
            TabIndex        =   54
            Top             =   750
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorUnitCIFNaMoeda 
            Height          =   300
            Left            =   4890
            TabIndex        =   55
            Top             =   480
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorUnitCIFEmReal 
            Height          =   300
            Left            =   2880
            TabIndex        =   56
            Top             =   495
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorTotalFOBNaMoeda 
            Height          =   300
            Left            =   2790
            TabIndex        =   57
            Top             =   1155
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorTotalFOBEmReal 
            Height          =   300
            Left            =   1335
            TabIndex        =   58
            Top             =   1140
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorTotalCIFNaMoeda 
            Height          =   225
            Left            =   1350
            TabIndex        =   59
            Top             =   765
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorTotalCIFEmReal 
            Height          =   300
            Left            =   2775
            TabIndex        =   60
            Top             =   810
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ItemPesoBruto 
            Height          =   300
            Left            =   1095
            TabIndex        =   99
            Top             =   1740
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ItemPesoLiq 
            Height          =   300
            Left            =   2655
            TabIndex        =   101
            Top             =   1755
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   3885
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   8820
            _ExtentX        =   15558
            _ExtentY        =   6853
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label DescDet 
            BorderStyle     =   1  'Fixed Single
            Height          =   570
            Left            =   135
            TabIndex        =   217
            Top             =   4230
            Width           =   8835
         End
      End
   End
   Begin VB.Frame FrameOpcao 
      BorderStyle     =   0  'None
      Height          =   5985
      Index           =   2
      Left            =   90
      TabIndex        =   42
      Top             =   645
      Visible         =   0   'False
      Width           =   9450
      Begin VB.CommandButton BotaoIPICodigo 
         Caption         =   "Classificação Fiscal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   75
         TabIndex        =   103
         Top             =   5280
         Width           =   1785
      End
      Begin VB.Frame FrameGridAdicao 
         Caption         =   "Adições"
         Height          =   5175
         Left            =   60
         TabIndex        =   43
         Top             =   60
         Width           =   9345
         Begin MSMask.MaskEdBox NumDrawback 
            Height          =   300
            Left            =   5460
            TabIndex        =   216
            Top             =   1875
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   11
            Mask            =   "###########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ICMSPercRedBase 
            Height          =   300
            Left            =   4530
            TabIndex        =   194
            Top             =   1050
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin VB.TextBox CodFabricante 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   5550
            TabIndex        =   192
            Top             =   2535
            Width           =   2250
         End
         Begin MSMask.MaskEdBox TaxaSiscomex 
            Height          =   300
            Left            =   6465
            TabIndex        =   191
            Top             =   3225
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DespAdua 
            Height          =   300
            Left            =   4650
            TabIndex        =   190
            Top             =   3630
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin VB.TextBox IPIDescricao 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   2490
            TabIndex        =   109
            Top             =   1380
            Width           =   1860
         End
         Begin MSMask.MaskEdBox IPICodigo 
            Height          =   300
            Left            =   900
            TabIndex        =   44
            Top             =   825
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox IPIAliquota 
            Height          =   300
            Left            =   1680
            TabIndex        =   45
            Top             =   900
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ICMSAliquota 
            Height          =   300
            Left            =   4380
            TabIndex        =   46
            Top             =   2520
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox IIAliquota 
            Height          =   300
            Left            =   2475
            TabIndex        =   47
            Top             =   465
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PISAliquota 
            Height          =   300
            Left            =   930
            TabIndex        =   70
            Top             =   2565
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox COFINSAliquota 
            Height          =   300
            Left            =   2745
            TabIndex        =   71
            Top             =   2595
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox IPIValor 
            Height          =   300
            Left            =   3225
            TabIndex        =   72
            Top             =   855
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox COFINSValor 
            Height          =   300
            Left            =   3060
            TabIndex        =   73
            Top             =   1740
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox IIValor 
            Height          =   300
            Left            =   3975
            TabIndex        =   74
            Top             =   450
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ICMSValor 
            Height          =   300
            Left            =   5430
            TabIndex        =   75
            Top             =   1185
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PISValor 
            Height          =   300
            Left            =   1305
            TabIndex        =   76
            Top             =   1710
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AdicaoValorAduaneiro 
            Height          =   300
            Left            =   900
            TabIndex        =   77
            Top             =   465
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PISBase 
            Height          =   300
            Left            =   1770
            TabIndex        =   78
            Top             =   4470
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ICMSBase 
            Height          =   300
            Left            =   4815
            TabIndex        =   79
            Top             =   4425
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox IPIBase 
            Height          =   300
            Left            =   135
            TabIndex        =   80
            Top             =   900
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox COFINSBase 
            Height          =   300
            Left            =   3060
            TabIndex        =   81
            Top             =   4425
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridAdicao 
            Height          =   3630
            Left            =   45
            TabIndex        =   48
            Top             =   285
            Width           =   9240
            _ExtentX        =   16298
            _ExtentY        =   6403
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame FrameOpcao 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   5475
      Index           =   4
      Left            =   165
      TabIndex        =   82
      Top             =   690
      Visible         =   0   'False
      Width           =   9240
      Begin VB.CommandButton BotaoTipoDespesa 
         Caption         =   "Tipos de Despesas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   180
         TabIndex        =   107
         Top             =   4515
         Width           =   1785
      End
      Begin VB.Frame Frame4 
         Caption         =   "Despesas"
         Height          =   4305
         Left            =   180
         TabIndex        =   83
         Top             =   120
         Width           =   8880
         Begin MSMask.MaskEdBox ComplDias 
            Height          =   300
            Left            =   7110
            TabIndex        =   165
            Top             =   2700
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   529
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
            Format          =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ComplPerc 
            Height          =   300
            Left            =   6210
            TabIndex        =   164
            Top             =   2700
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   529
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
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin VB.TextBox ComplDescricao 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   885
            TabIndex        =   94
            Top             =   3420
            Width           =   4005
         End
         Begin VB.TextBox ComplTipo 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   255
            TabIndex        =   93
            Top             =   3390
            Width           =   585
         End
         Begin MSMask.MaskEdBox ComplValor 
            Height          =   300
            Left            =   4950
            TabIndex        =   95
            Top             =   3435
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
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
         Begin MSFlexGridLib.MSFlexGrid GridDespesas 
            Height          =   2880
            Left            =   150
            TabIndex        =   84
            Top             =   285
            Width           =   8520
            _ExtentX        =   15028
            _ExtentY        =   5080
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Label LabelOutrasDesp 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5820
         TabIndex        =   169
         Top             =   4725
         Width           =   1170
      End
      Begin VB.Label Label36 
         Caption         =   "Outras Despesas:"
         Height          =   480
         Left            =   4920
         TabIndex        =   168
         Top             =   4725
         Width           =   810
      End
      Begin VB.Label LabelDespICMS 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3435
         TabIndex        =   167
         Top             =   4740
         Width           =   1230
      End
      Begin VB.Label Label30 
         Caption         =   "Despesas na base do ICMS:"
         Height          =   465
         Left            =   2310
         TabIndex        =   166
         Top             =   4650
         Width           =   1080
      End
   End
   Begin VB.Frame FrameOpcao 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   5250
      Index           =   5
      Left            =   135
      TabIndex        =   85
      Top             =   705
      Visible         =   0   'False
      Width           =   9180
      Begin VB.Frame Frame5 
         Caption         =   "Resumo da DI"
         Height          =   4905
         Left            =   150
         TabIndex        =   86
         Top             =   105
         Width           =   8790
         Begin MSMask.MaskEdBox DIIIValor 
            Height          =   330
            Left            =   1950
            TabIndex        =   100
            Top             =   2220
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DIIPIValor 
            Height          =   330
            Left            =   1950
            TabIndex        =   102
            Top             =   2646
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DIPISValor 
            Height          =   330
            Left            =   1950
            TabIndex        =   104
            Top             =   3072
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DICOFINSValor 
            Height          =   330
            Left            =   1950
            TabIndex        =   106
            Top             =   3498
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DIICMSValor 
            Height          =   330
            Left            =   1950
            TabIndex        =   108
            Top             =   4350
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DIValorDespesas 
            Height          =   330
            Left            =   1950
            TabIndex        =   98
            Top             =   3924
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DIValorFOB 
            Height          =   330
            Left            =   1965
            TabIndex        =   132
            Top             =   360
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DIValorFrete 
            Height          =   330
            Left            =   1965
            TabIndex        =   134
            Top             =   750
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DIValorSeguro 
            Height          =   330
            Left            =   1965
            TabIndex        =   136
            Top             =   1125
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DIValorCIF 
            Height          =   330
            Left            =   1965
            TabIndex        =   138
            Top             =   1530
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DIValorProdutos 
            Height          =   330
            Left            =   6270
            TabIndex        =   140
            Top             =   375
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DIValorOutrasDesp 
            Height          =   330
            Left            =   6270
            TabIndex        =   142
            Top             =   825
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DIValorTotal 
            Height          =   330
            Left            =   6270
            TabIndex        =   144
            Top             =   1275
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
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
            Height          =   240
            Left            =   4665
            TabIndex        =   145
            ToolTipText     =   "Imposto de Importação"
            Top             =   1320
            Width           =   1470
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "Outras Desp.:"
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
            Left            =   4650
            TabIndex        =   143
            ToolTipText     =   "Imposto de Importação"
            Top             =   885
            Width           =   1470
         End
         Begin VB.Label LabelDIValorProdutos 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor Produtos:"
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
            Left            =   4665
            TabIndex        =   141
            ToolTipText     =   "Imposto de Importação"
            Top             =   420
            Width           =   1470
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor CIF:"
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
            Left            =   825
            TabIndex        =   139
            ToolTipText     =   "Imposto de Importação"
            Top             =   1575
            Width           =   1005
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Seguro:"
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
            Left            =   825
            TabIndex        =   137
            ToolTipText     =   "Imposto de Importação"
            Top             =   1170
            Width           =   1005
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Frete:"
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
            Left            =   825
            TabIndex        =   135
            ToolTipText     =   "Imposto de Importação"
            Top             =   795
            Width           =   1005
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor FOB:"
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
            Left            =   825
            TabIndex        =   133
            ToolTipText     =   "Imposto de Importação"
            Top             =   405
            Width           =   1005
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Despesas:"
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
            Left            =   810
            TabIndex        =   92
            Top             =   3981
            Width           =   1005
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "II:"
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
            Left            =   810
            TabIndex        =   91
            ToolTipText     =   "Imposto de Importação"
            Top             =   2265
            Width           =   1005
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "IPI:"
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
            Left            =   810
            TabIndex        =   90
            Top             =   2694
            Width           =   1005
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "PIS:"
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
            Left            =   810
            TabIndex        =   89
            Top             =   3123
            Width           =   1005
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "COFINS:"
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
            Left            =   810
            TabIndex        =   88
            Top             =   3552
            Width           =   1005
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "ICMS:"
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
            Left            =   810
            TabIndex        =   87
            Top             =   4410
            Width           =   1005
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   6960
      ScaleHeight     =   450
      ScaleWidth      =   2550
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   30
      Width           =   2610
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   60
         Picture         =   "DIInfo.ctx":0185
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Imprimir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   555
         Picture         =   "DIInfo.ctx":0287
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1065
         Picture         =   "DIInfo.ctx":03E1
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1560
         Picture         =   "DIInfo.ctx":056B
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2040
         Picture         =   "DIInfo.ctx":0A9D
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   6510
      Left            =   60
      TabIndex        =   62
      Top             =   270
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   11483
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Adições"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Despesas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Resumo"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pedido Compra"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label23 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   6420
      TabIndex        =   97
      Top             =   1650
      Width           =   1320
   End
   Begin VB.Label Label21 
      Caption         =   "Mercadoria:"
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
      Left            =   5295
      TabIndex        =   96
      Top             =   1665
      Width           =   1005
   End
End
Attribute VB_Name = "DIInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const NUM_MAX_LINHAS_GRID_DESPESAS = 50
Private Const NUM_MAX_LINHAS_GRID_ADICAO = 200
Private Const NUM_MAX_LINHAS_GRID_ITENS = 990

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gbTrazendoPC As Boolean

Private iDataAlterada As Integer
Private sNumDIAnt As String
Private gdtDataAnterior As Date

Private iTaxaMoeda1Alterada As Integer
Private iTaxaMoeda2Alterada As Integer

Private iMercEmMoedaAlterada As Integer
Private iFreteEmMoedaAlterada As Integer
Private iSeguroEmMoedaAlterada As Integer

Private iMoedaMercadoriaAnt As Integer
Private iMoedaFreteAnt As Integer
Private iMoedaSeguroAnt As Integer
Private iMoedaItensAnt As Integer

Private iMercEmRealAlterada As Integer
Private iFreteEmRealAlterada As Integer
Private iSeguroEmRealAlterada As Integer

Private iDIPesoBrutoAlterada As Integer
Private iDIPesoLiqAlterada As Integer

Dim iAlterado As Integer
Dim iFornecedorAlterado As Integer
Dim iFrameAtual As Integer

Dim objGridDespesas As AdmGrid
Dim iGrid_ComplTipo_Col As Integer
Dim iGrid_ComplDescricao_Col As Integer
Dim iGrid_ComplValor_Col As Integer
Dim iGrid_ComplPerc_Col As Integer
Dim iGrid_ComplDias_Col As Integer

Dim objGridAdicao As AdmGrid
Dim iGrid_IPICodigo_Col As Integer
Dim iGrid_IPIDescricao_Col As Integer
Dim iGrid_AdicaoValorAduaneiro_Col As Integer
Dim iGrid_IIAliquota_Col As Integer
Dim iGrid_IIValor_Col As Integer
Dim iGrid_IPIBase_Col As Integer
Dim iGrid_IPIAliquota_Col As Integer
Dim iGrid_IPIValor_Col As Integer
Dim iGrid_PISBase_Col As Integer
Dim iGrid_PISAliquota_Col As Integer
Dim iGrid_PISValor_Col As Integer
Dim iGrid_COFINSBase_Col As Integer
Dim iGrid_COFINSAliquota_Col As Integer
Dim iGrid_COFINSValor_Col As Integer
Dim iGrid_ICMSBase_Col As Integer
Dim iGrid_ICMSAliquota_Col As Integer
Dim iGrid_ICMSValor_Col As Integer
Dim iGrid_DespAdua_Col As Integer
Dim iGrid_TaxaSiscomex_Col As Integer
Dim iGrid_CodFabricante_Col As Integer
Dim iGrid_ICMSPercRedBase_Col As Integer
Dim iGrid_NumDrawback_Col As Integer

Dim objGridItens As AdmGrid
Dim iGrid_AdicaoItem_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_UnidadeMed_Col As Integer
Dim iGrid_ValorUnitFOBNaMoeda_Col As Integer
Dim iGrid_ValorUnitFOBEmReal_Col As Integer
Dim iGrid_ValorUnitCIFNaMoeda_Col As Integer
Dim iGrid_ValorUnitCIFEmReal_Col As Integer
Dim iGrid_ValorTotalFOBNaMoeda_Col As Integer
Dim iGrid_ValorTotalFOBEmReal_Col As Integer
Dim iGrid_ValorTotalCIFNaMoeda_Col As Integer
Dim iGrid_ValorTotalCIFEmReal_Col As Integer
Dim iGrid_TotalCIFEmRealManual_Col As Integer
Dim iGrid_ItemPesoBruto_Col As Integer
Dim iGrid_ItemPesoLiq_Col As Integer
Dim iGrid_IPIValorUnitario_Col As Integer

'GridItensPC
Public objGridItensPC As AdmGrid
Public iGrid_CodigoPC_Col As Integer
Public iGrid_DataPC_Col As Integer
Public iGrid_ProdutoPC_Col As Integer
Public iGrid_DescProdutoPC_Col As Integer
Public iGrid_UMPC_Col As Integer
Public iGrid_QuantPC_Col As Integer


Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoTipoImport As AdmEvento
Attribute objEventoTipoImport.VB_VarHelpID = -1
Private WithEvents objEventoIPICodigo As AdmEvento
Attribute objEventoIPICodigo.VB_VarHelpID = -1
Private WithEvents objEventoPC As AdmEvento
Attribute objEventoPC.VB_VarHelpID = -1
Private WithEvents objEventoItemPC As AdmEvento
Attribute objEventoItemPC.VB_VarHelpID = -1

Private gbCarregandoTela As Boolean

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Declaração de importação"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "DIInfo"

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

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objDIInfo As New ClassDIInfo
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_BotaoImprimir_Click

    If Len(Trim(Numero.Text)) = 0 Then gError 202737
    
    objDIInfo.sNumero = Numero.Text
    
    lErro = CF("DIInfo_Le", objDIInfo)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 202738
    
    If lErro <> SUCESSO Then gError 202739 'DI NÃO CADASTRADA
    
    'Imprime o Relatorio de OP
    lErro = objRelatorio.ExecutarDireto("Declaração de Importação", "Numero = @TNUMERO", 0, "DI", "TNUMERO", objDIInfo.sNumero)
    If lErro <> SUCESSO Then gError 202740
    
    Exit Sub
    
Erro_BotaoImprimir_Click:

    Select Case gErr
    
        Case 202737
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DIINFO_NAO_PREENCHIDO", gErr)
            
        Case 202738, 202740
        
        Case 202739
            Call Rotina_Erro(vbOKOnly, "ERRO_DIINFO_NAO_CADASTRADO", gErr, objDIInfo.sNumero)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 202741)

    End Select
    
End Sub


Private Sub BotaoLimparGridPC_Click()
    
    Call Grid_Limpa(objGridItensPC)

End Sub

Private Sub TotalCIFEmRealManual_Click()

    If gbCarregandoTela = False Then

        'Registra que houve alteração
        iAlterado = REGISTRO_ALTERADO

        Call Calcula_ValorCIF_Linha(GridItens.Row)
        Call Calcula_Valores_Adicao(StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_AdicaoItem_Col)), 0)
    
    End If
    
End Sub

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

    iFrameAtual = 1

    Set objGridDespesas = Nothing
    Set objGridAdicao = Nothing
    Set objGridItens = Nothing
    Set objGridItensPC = Nothing

    Set objEventoNumero = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoProduto = Nothing
    Set objEventoTipoImport = Nothing
    Set objEventoIPICodigo = Nothing
    Set objEventoPC = Nothing
    Set objEventoItemPC = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196561)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long
Dim colCodigo As New Collection
Dim vCodigo As Variant

On Error GoTo Erro_Form_Load

    gbCarregandoTela = False
    
    Set objEventoNumero = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoTipoImport = New AdmEvento
    Set objEventoIPICodigo = New AdmEvento
    Set objEventoPC = New AdmEvento
    Set objEventoItemPC = New AdmEvento
    
    gbTrazendoPC = False

    'carrega a combo de Moedas
    lErro = Carrega_Moeda() 'leo
    If lErro <> SUCESSO Then gError 196562
    
'    TaxaMoeda.Format = FORMATO_TAXA_CONVERSAO_MOEDA
    
    lErro = Inicializa_GridAdicao(objGridAdicao)
    If lErro <> SUCESSO Then gError 196563

    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError 196564

    lErro = Inicializa_GridDespesas(objGridDespesas)
    If lErro <> SUCESSO Then gError 196565
    
    lErro = Inicializa_GridItensPC(objGridItensPC)
    If lErro <> SUCESSO Then gError 210566
    
    'Inicializa a Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 196566
    
    'Inicializa a Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoPC)
    If lErro <> SUCESSO Then gError 210578
    
    'Lê cada codigo da tabela Estados
    lErro = CF("Codigos_Le", "Estados", "Sigla", TIPO_STR, colCodigo, STRING_ESTADOS_SIGLA)
    If lErro <> SUCESSO Then gError 202958

    'Preenche as ComboBox Estados com os objetos da colecao colCodigo
    DUF.AddItem ""
    UFAdquir.AddItem ""
    For Each vCodigo In colCodigo
        DUF.AddItem vCodigo
        UFAdquir.AddItem vCodigo
    Next
    
    MoedaFrete.ListIndex = 0
    MoedaItens.ListIndex = 0
    MoedaMercadoria.ListIndex = 0
    MoedaSeguro.ListIndex = 0

    Intermedio.ListIndex = 0
    ViaTransp.ListIndex = 0

    iFrameAtual = 1

    Call ZerarFlagsAlteracao
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 196562 To 196566, 210566, 210578

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196567)

    End Select

    Call ZerarFlagsAlteracao
    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objDIInfo As ClassDIInfo) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objDIInfo Is Nothing) Then

        lErro = Traz_DIInfo_Tela(objDIInfo)
        If lErro <> SUCESSO Then gError 196568

    End If

    Call ZerarFlagsAlteracao
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 196568

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196569)

    End Select

    Call ZerarFlagsAlteracao
    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objDIInfo As ClassDIInfo, Optional ByVal bMovTudo As Boolean = True) As Long

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Move_Tela_Memoria

    objDIInfo.sNumero = Trim(Numero.Text)
    objDIInfo.dtData = StrParaDate(Data.Text)
    objDIInfo.iFilialEmpresa = giFilialEmpresa
    objDIInfo.sDescricao = DIDescricao.Text
    
    'Verifica preenchimento de Fornecedor
    If Len(Trim(Fornecedor.Text)) <> 0 Then

        objFornecedor.sNomeReduzido = Fornecedor.Text

        'Lê Fornecedor no BD
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 196570

        'Se não achou o Fornecedor --> erro
        If lErro = 6681 Then gError 196571

        objDIInfo.lFornTrading = objFornecedor.lCodigo

    End If

    objDIInfo.iFilialFornTrading = Codigo_Extrai(Filial.Text)
        
    objDIInfo.sProcessoTrading = ProcessoTrading.Text
    objDIInfo.iMoeda1 = Codigo_Extrai(Moeda1.Text)
    objDIInfo.dTaxaMoeda1 = StrParaDbl(TaxaMoeda1.Text)
    objDIInfo.iMoeda2 = Codigo_Extrai(Moeda2.Text)
    objDIInfo.dTaxaMoeda2 = StrParaDbl(TaxaMoeda2.Text)
    
    objDIInfo.sUFDesembaraco = DUF.Text
    objDIInfo.sLocalDesembaraco = DLocal.Text
    objDIInfo.dtDataDesembaraco = StrParaDate(DData.Text)
    objDIInfo.iMoedaFrete = StrParaInt(MoedaFrete.Text)
    objDIInfo.iMoedaItens = StrParaInt(MoedaItens.Text)
    objDIInfo.iMoedaMercadoria = StrParaInt(MoedaMercadoria.Text)
    objDIInfo.iMoedaSeguro = StrParaInt(MoedaSeguro.Text)
    
    objDIInfo.sCodExportador = CodExportador.Text
    
    objDIInfo.dPesoBrutoKG = StrParaDbl(DIPesoBruto.Text)
    objDIInfo.dPesoLiqKG = StrParaDbl(DIPesoLiquido.Text)
    objDIInfo.dValorMercadoriaMoeda = StrParaDbl(ValorMercadoriaMoeda.Text)
    objDIInfo.dValorFreteInternacMoeda = StrParaDbl(ValorFreteInternacMoeda.Text)
    objDIInfo.dValorSeguroInternacMoeda = StrParaDbl(ValorSeguroInternacMoeda.Text)
    objDIInfo.dValorMercadoriaEmReal = StrParaDbl(ValorMercadoriaEmReal.Text)
    objDIInfo.dValorFreteInternacEmReal = StrParaDbl(ValorFreteInternacEmReal.Text)
    objDIInfo.dValorSeguroInternacEmReal = StrParaDbl(ValorSeguroInternacEmReal.Text)
    
    'do tab de resumo
    objDIInfo.dValorDespesas = StrParaDbl(DIValorDespesas.Text)
    objDIInfo.dIIValor = StrParaDbl(DIIIValor.Text)
    objDIInfo.dIPIValor = StrParaDbl(DIIPIValor.Text)
    objDIInfo.dPISValor = StrParaDbl(DIPISValor.Text)
    objDIInfo.dCOFINSValor = StrParaDbl(DICOFINSValor.Text)
    objDIInfo.dICMSValor = StrParaDbl(DIICMSValor.Text)
    
    objDIInfo.iIntermedio = Codigo_Extrai(Intermedio.Text)
    objDIInfo.iViaTransp = Codigo_Extrai(ViaTransp.Text)
    objDIInfo.sUFAdquir = Trim(UFAdquir.Text)
    objDIInfo.sCNPJAdquir = Trim(CNPJAdquir.ClipText)
    
    If bMovTudo Then
            
        lErro = Move_GridAdicao_Memoria(objDIInfo)
        If lErro <> SUCESSO Then gError 196573
        
        lErro = Move_GridDespesas_Memoria(objDIInfo)
        If lErro <> SUCESSO Then gError 196574
        
        lErro = Move_GridItensPC_Memoria(objDIInfo)
        If lErro <> SUCESSO Then gError 210593
        
    End If
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 196570, 196573, 196574, 210593

        Case 196571
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196575)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objDIInfo As New ClassDIInfo

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "DIInfo"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objDIInfo, False)
    If lErro <> SUCESSO Then gError 196576

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Numero", objDIInfo.sNumero, STRING_DI_NUMERO, "Numero"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 196576

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196577)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objDIInfo As New ClassDIInfo

On Error GoTo Erro_Tela_Preenche

    objDIInfo.sNumero = colCampoValor.Item("Numero").vValor

    If Len(Trim(objDIInfo.sNumero)) > 0 Then

        lErro = Traz_DIInfo_Tela(objDIInfo)
        If lErro <> SUCESSO Then gError 196578

    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 196578

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196579)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objDIInfo As New ClassDIInfo, sNumDI As String

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Numero.Text)) = 0 Then gError 196580
    sNumDI = Replace(Replace(Trim(Numero.Text), "/", ""), "-", "")
    If Len(sNumDI) > 12 Then gError 201179
    
    If Len(Trim(Data.ClipText)) = 0 Then gError 196581
    If Len(Trim(Moeda1.Text)) = 0 Then gError 196582
    If Len(Trim(TaxaMoeda1.Text)) = 0 Then gError 196583
    If Len(Trim(CodExportador.Text)) = 0 Then gError 206781
    If Len(Trim(DUF.Text)) = 0 Then gError 206782
    If Len(Trim(DLocal.Text)) = 0 Then gError 206783
    If Len(Trim(DData.ClipText)) = 0 Then gError 206784
    '#####################
    
    If StrParaDbl(DIPesoBruto.Text) < StrParaDbl(DIPesoLiquido.Text) Then gError 196584
    
    If Codigo_Extrai(Intermedio.Text) <> 1 Then
        If Len(Trim(CNPJAdquir.Text)) = 0 Then gError 213587
        If Len(Trim(UFAdquir.Text)) = 0 Then gError 213588
    End If

    'Preenche o objDIInfo
    lErro = Move_Tela_Memoria(objDIInfo)
    If lErro <> SUCESSO Then gError 196585

    'criticas:
    lErro = DIInfo_Critica(objDIInfo)
    If lErro <> SUCESSO Then gError 196585

    lErro = Trata_Alteracao(objDIInfo, objDIInfo.sNumero)
    If lErro <> SUCESSO Then gError 196586
    
    'Grava o/a DIInfo no Banco de Dados
    lErro = CF("DIInfo_Grava", objDIInfo)
    If lErro <> SUCESSO Then gError 196587

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 201179
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DIINFO_FORMATO", gErr)
        
        Case 196580
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DIINFO_NAO_PREENCHIDO", gErr)

        Case 196581 'ERRO_DATA_NAO_PREENCHIDA
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case 196582
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_NAO_PREENCHIDA", gErr)
            
        Case 196583 'ERRO_TAXA_GRID_IMCOMPLETA
            Call Rotina_Erro(vbOKOnly, "ERRO_TAXA_GRID_IMCOMPLETA", gErr)
        
        Case 196584 'ERRO_PESO_LIQUIDO_MAIOR_BRUTO
            Call Rotina_Erro(vbOKOnly, "ERRO_PESO_LIQUIDO_MAIOR_BRUTO", gErr, DIPesoLiquido.Text, DIPesoBruto.Text)
        
        Case 196585 To 196587
        
        Case 206781
            Call Rotina_Erro(vbOKOnly, "ERRO_CODEXPORTADOR_NAO_PREENCHIDO", gErr)
        
        Case 206782
            Call Rotina_Erro(vbOKOnly, "ERRO_UFDESEMBARACO_NAO_PREENCHIDA", gErr)
        
        Case 206783
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCALDESEMBARACO_NAO_PREENCHIDO", gErr)
        
        Case 206784
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADESEMBARACO_NAO_PREENCHIDA", gErr)
        
        Case 213587, 213588
            Call Rotina_Erro(vbOKOnly, "ERRO_DI_NAO_IMP_PROPIA_SEM_CNPJ_UF", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196588)

    End Select

    Exit Function

End Function

Function Limpa_Tela_DIInfo() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_DIInfo

    Moeda1.ListIndex = -1
    Moeda2.ListIndex = -1

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    sNumDIAnt = ""
    
    gdtDataAnterior = DATA_NULA

    LabelTotalFOBAdicaoMoeda.Caption = ""
    LabelTotalFOBAdicaoReal.Caption = ""
    LabelTotalCIFAdicaoMoeda.Caption = ""
    LabelTotalCIFAdicaoReal.Caption = ""
    
    LabelTotalFOBDIMoeda.Caption = ""
    LabelTotalFOBDIReal.Caption = ""
    LabelTotalCIFDIMoeda.Caption = ""
    LabelTotalCIFDIReal.Caption = ""
    LabelPesoLiqAdicao.Caption = ""
    LabelPesoLiqDI.Caption = ""

    Call Grid_Limpa(objGridItens)
    Call Grid_Limpa(objGridAdicao)
    Call Grid_Limpa(objGridDespesas)
    Call Grid_Limpa(objGridItensPC)
    
    LabelDespICMS.Caption = ""
    LabelOutrasDesp.Caption = ""
    
    DUF.ListIndex = -1
    MoedaFrete.ListIndex = 0
    MoedaItens.ListIndex = 0
    MoedaMercadoria.ListIndex = 0
    MoedaSeguro.ListIndex = 0
    
    Intermedio.ListIndex = 0
    ViaTransp.ListIndex = 0
       
    AdicaoItem.Clear
    
    Call Limpa_PC

    Call ZerarFlagsAlteracao
    iAlterado = 0

    Limpa_Tela_DIInfo = SUCESSO

    Exit Function

Erro_Limpa_Tela_DIInfo:

    Limpa_Tela_DIInfo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196589)

    End Select

    Exit Function

End Function

Function Traz_DIInfo_Tela(objDIInfo As ClassDIInfo) As Long

Dim lErro As Long, iIndice As Integer

On Error GoTo Erro_Traz_DIInfo_Tela

    gbCarregandoTela = True
    
    Call Limpa_Tela_DIInfo
    
    sNumDIAnt = objDIInfo.sNumero
    Numero.Text = objDIInfo.sNumero

    'Lê o DIInfo que está sendo Passado
    lErro = CF("DIInfo_Le", objDIInfo)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 196590

    If lErro = SUCESSO Then

        Numero.Text = objDIInfo.sNumero

        Call DateParaMasked(Data, objDIInfo.dtData)
        gdtDataAnterior = objDIInfo.dtData

        '??? completar
'''        If objDIInfo.iStatus <> 0 Then
'''            Status.PromptInclude = False
'''            Status.Text = CStr(objDIInfo.iStatus)
'''            Status.PromptInclude = True
'''        End If

        DIDescricao.Text = objDIInfo.sDescricao

        If objDIInfo.lFornTrading <> 0 Then
        
            'preenche o Fornecedor
            Fornecedor.Text = objDIInfo.lFornTrading
            Call Fornecedor_Validate(bSGECancelDummy)
        
            'Preenche a Filial do Fornecedor
            Filial.Text = objDIInfo.iFilialFornTrading
            Call Filial_Validate(bSGECancelDummy)
        
        End If
        
        ProcessoTrading.Text = objDIInfo.sProcessoTrading

        For iIndice = 0 To Moeda1.ListCount - 1
            If objDIInfo.iMoeda1 = Codigo_Extrai(Moeda1.List(iIndice)) Then
                Moeda1.ListIndex = iIndice
                Exit For
            End If
        Next

        If objDIInfo.dTaxaMoeda1 <> 0 Then
            TaxaMoeda1.PromptInclude = False
            TaxaMoeda1.Text = Format(objDIInfo.dTaxaMoeda1, TaxaMoeda1.Format)
            TaxaMoeda1.PromptInclude = True
        End If
        
        For iIndice = 0 To Moeda2.ListCount - 1
            If objDIInfo.iMoeda2 = Codigo_Extrai(Moeda2.List(iIndice)) Then
                Moeda2.ListIndex = iIndice
                Exit For
            End If
        Next

        If objDIInfo.dTaxaMoeda2 <> 0 Then
            TaxaMoeda2.PromptInclude = False
            TaxaMoeda2.Text = Format(objDIInfo.dTaxaMoeda2, TaxaMoeda2.Format)
            TaxaMoeda2.PromptInclude = True
        End If
        
        For iIndice = 0 To DUF.ListCount - 1
            If objDIInfo.sUFDesembaraco = DUF.List(iIndice) Then
                DUF.ListIndex = iIndice
                Exit For
            End If
        Next
    
        DLocal.Text = objDIInfo.sLocalDesembaraco
        
        If objDIInfo.dtDataDesembaraco <> DATA_NULA Then
            DData.PromptInclude = False
            DData.Text = Format(objDIInfo.dtDataDesembaraco, "dd/mm/yy")
            DData.PromptInclude = True
        End If
        
        MoedaFrete.ListIndex = objDIInfo.iMoedaFrete - 1
        MoedaItens.ListIndex = objDIInfo.iMoedaItens - 1
        MoedaMercadoria.ListIndex = objDIInfo.iMoedaMercadoria - 1
        MoedaSeguro.ListIndex = objDIInfo.iMoedaSeguro - 1
        
        CodExportador.Text = objDIInfo.sCodExportador
        
        If objDIInfo.dPesoBrutoKG <> 0 Then
            DIPesoBruto.Text = Formata_Estoque(objDIInfo.dPesoBrutoKG)
        End If

        If objDIInfo.dPesoLiqKG <> 0 Then
            DIPesoLiquido.Text = Formata_Estoque(objDIInfo.dPesoLiqKG)
        End If
        
        Call Combo_Seleciona_ItemData(Intermedio, objDIInfo.iIntermedio)
        Call Combo_Seleciona_ItemData(ViaTransp, objDIInfo.iViaTransp)
        For iIndice = 0 To UFAdquir.ListCount - 1
            If objDIInfo.sUFAdquir = UFAdquir.List(iIndice) Then
                UFAdquir.ListIndex = iIndice
                Exit For
            End If
        Next
        If objDIInfo.sCNPJAdquir <> "" Then
            CNPJAdquir.Text = objDIInfo.sCNPJAdquir
            Call CNPJAdquir_Validate(bSGECancelDummy)
        Else
            CNPJAdquir.Text = ""
        End If

        If objDIInfo.dValorMercadoriaMoeda <> 0 Then
            ValorMercadoriaMoeda.PromptInclude = False
            ValorMercadoriaMoeda.Text = Format(objDIInfo.dValorMercadoriaMoeda, ValorMercadoriaMoeda.Format)
            ValorMercadoriaMoeda.PromptInclude = True
        End If


        If objDIInfo.dValorFreteInternacMoeda <> 0 Then
            ValorFreteInternacMoeda.PromptInclude = False
            ValorFreteInternacMoeda.Text = Format(objDIInfo.dValorFreteInternacMoeda, ValorFreteInternacMoeda.Format)
            ValorFreteInternacMoeda.PromptInclude = True
        End If


        If objDIInfo.dValorSeguroInternacMoeda <> 0 Then
            ValorSeguroInternacMoeda.PromptInclude = False
            ValorSeguroInternacMoeda.Text = Format(objDIInfo.dValorSeguroInternacMoeda, ValorSeguroInternacMoeda.Format)
            ValorSeguroInternacMoeda.PromptInclude = True
        End If


        If objDIInfo.dValorMercadoriaEmReal <> 0 Then
            ValorMercadoriaEmReal.PromptInclude = False
            ValorMercadoriaEmReal.Text = Format(objDIInfo.dValorMercadoriaEmReal, ValorMercadoriaEmReal.Format)
            ValorMercadoriaEmReal.PromptInclude = True
        End If


        If objDIInfo.dValorFreteInternacEmReal <> 0 Then
            ValorFreteInternacEmReal.PromptInclude = False
            ValorFreteInternacEmReal.Text = Format(objDIInfo.dValorFreteInternacEmReal, ValorFreteInternacEmReal.Format)
            ValorFreteInternacEmReal.PromptInclude = True
        End If


        If objDIInfo.dValorSeguroInternacEmReal <> 0 Then
            ValorSeguroInternacEmReal.PromptInclude = False
            ValorSeguroInternacEmReal.Text = Format(objDIInfo.dValorSeguroInternacEmReal, ValorSeguroInternacEmReal.Format)
            ValorSeguroInternacEmReal.PromptInclude = True
        End If

'        'do tab de resumo
'        If objDIInfo.dValorDespesas <> 0 Then
'            DIValorDespesas.PromptInclude = False
'            DIValorDespesas.Text = Format(objDIInfo.dValorDespesas, DIValorDespesas.Format)
'            DIValorDespesas.PromptInclude = True
'        End If
'
'        If objDIInfo.dIIValor <> 0 Then
'            DIIIValor.PromptInclude = False
'            DIIIValor.Text = Format(objDIInfo.dIIValor, DIIIValor.Format)
'            DIIIValor.PromptInclude = True
'        End If
'
'        If objDIInfo.dIPIValor <> 0 Then
'            DIIPIValor.PromptInclude = False
'            DIIPIValor.Text = Format(objDIInfo.dIPIValor, DIIPIValor.Format)
'            DIIPIValor.PromptInclude = True
'        End If
'
'        If objDIInfo.dPISValor <> 0 Then
'            DIPISValor.PromptInclude = False
'            DIPISValor.Text = Format(objDIInfo.dPISValor, DIPISValor.Format)
'            DIPISValor.PromptInclude = True
'        End If
'
'        If objDIInfo.dCOFINSValor <> 0 Then
'            DICOFINSValor.PromptInclude = False
'            DICOFINSValor.Text = Format(objDIInfo.dCOFINSValor, DICOFINSValor.Format)
'            DICOFINSValor.PromptInclude = True
'        End If
'
'        If objDIInfo.dICMSValor <> 0 Then
'            DIICMSValor.PromptInclude = False
'            DIICMSValor.Text = Format(objDIInfo.dICMSValor, DIICMSValor.Format)
'            DIICMSValor.PromptInclude = True
'        End If

        lErro = Preenche_GridAdicao_Tela(objDIInfo)
        If lErro <> SUCESSO Then gError 196591
        
        lErro = Preenche_GridDespesas_Tela(objDIInfo)
        If lErro <> SUCESSO Then gError 196592
    
        lErro = Preenche_GridItensPCDI_Tela(objDIInfo)
        If lErro <> SUCESSO Then gError 210605
    
        Call Calcula_Valores
            
        Call ZerarFlagsAlteracao
        
        iAlterado = 0

    End If
    
    gbCarregandoTela = False
    
    Traz_DIInfo_Tela = SUCESSO

    Exit Function

Erro_Traz_DIInfo_Tela:

    Traz_DIInfo_Tela = gErr

    Select Case gErr

        Case 196590 To 196592, 210605

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196593)

    End Select

    gbCarregandoTela = False
    
    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 196594

    'Limpa Tela
    Call Limpa_Tela_DIInfo

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 196594

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196595)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196596)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 196597

    Call Limpa_Tela_DIInfo

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 196597

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196598)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objDIInfo As New ClassDIInfo
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Numero.Text)) = 0 Then gError 196599

    objDIInfo.sNumero = Numero.Text

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_DIINFO", objDIInfo.sNumero)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("DIInfo_Exclui", objDIInfo)
        If lErro <> SUCESSO Then gError 196600

        'Limpa Tela
        Call Limpa_Tela_DIInfo

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 196599
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DIINFO_NAO_PREENCHIDO", gErr)

        Case 196600

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196601)

    End Select

    Exit Sub

End Sub

Private Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objDI As New ClassDIInfo
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_Numero_Validate

    'Verifica se Numero está preenchida, é diferente do anterior e se não estão apenas trazendo algo gravado para tela
    If Len(Trim(Numero.Text)) <> 0 And sNumDIAnt <> Trim(Numero.Text) And Not gbCarregandoTela Then

        objDI.sNumero = Trim(Numero.Text)
        objDI.dtData = StrParaDate(Data.Text)
        objDI.iFilialEmpresa = giFilialEmpresa

        lErro = CF("DIInfo_Le", objDI)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
        If lErro = SUCESSO Then
        'DI Já cadastrada então traz para tela
        
            lErro = Traz_DIInfo_Tela(objDI)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        Else
        
            lErro = CF("DIInfo_Le_XML", objDI)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
            
            If lErro = SUCESSO Then
            'A DI não está cadastrada mas existe o XML dela importado, pergunta se deseja carregar os dados
            
                vbResult = Rotina_Aviso(vbYesNo, "AVISO_DI_COM_XML")
                If vbResult = vbYes Then
                
                    lErro = Traz_DIInfo_Tela_XML(objDI)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                               
                End If
            
            End If
        
        End If
        
    End If
    
    sNumDIAnt = Trim(Numero.Text)

    Exit Sub

Erro_Numero_Validate:

    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196602)

    End Select

    Exit Sub

End Sub

Private Sub Numero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 196603

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 196603

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196604)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_UpClick

    Data.SetFocus

    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 196605

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 196605

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196606)

    End Select

    Exit Sub

End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)
    
End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long, dtDataNova As Date

On Error GoTo Erro_Data_Validate

    If iDataAlterada <> 0 Then

        If Len(Trim(Data.ClipText)) <> 0 Then
    
            lErro = Data_Critica(Data.Text)
            If lErro <> SUCESSO Then gError 196607
    
        End If
        
        dtDataNova = StrParaDate(Data.Text)
        
        If gbCarregandoTela = False Then
        
            If (dtDataNova = DATA_NULA Or dtDataNova >= DATA_PIS_NOVO_CALC) And (gdtDataAnterior <> DATA_NULA And gdtDataAnterior < DATA_PIS_NOVO_CALC) Or _
                (gdtDataAnterior = DATA_NULA Or gdtDataAnterior >= DATA_PIS_NOVO_CALC) And (dtDataNova <> DATA_NULA And dtDataNova < DATA_PIS_NOVO_CALC) Then Call Calcula_Valores_Adicao(0, 0)
            
        End If
        
        gdtDataAnterior = dtDataNova
        
        iDataAlterada = 0

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 196607

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196608)

    End Select

    Exit Sub

End Sub

Private Sub Data_Change()
    iAlterado = REGISTRO_ALTERADO
    iDataAlterada = REGISTRO_ALTERADO
End Sub

Private Sub ProcessoTrading_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProcessoTrading_Validate

    'Verifica se ProcessoTrading está preenchida
    If Len(Trim(ProcessoTrading.Text)) <> 0 Then

       '#######################################
       'CRITICA ProcessoTrading
       '#######################################

    End If

    Exit Sub

Erro_ProcessoTrading_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184569)

    End Select

    Exit Sub

End Sub

Private Sub ProcessoTrading_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorMercadoriaMoeda_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dTaxa As Double

On Error GoTo Erro_ValorMercadoriaMoeda_Validate

    If iMercEmMoedaAlterada <> 0 Or iMoedaMercadoriaAnt <> StrParaInt(MoedaMercadoria.Text) Then
    
        iMoedaMercadoriaAnt = StrParaInt(MoedaMercadoria.Text)
    
        'Verifica se ValorMercadoriaMoeda está preenchida
        If Len(Trim(ValorMercadoriaMoeda.Text)) <> 0 Then
    
           'Critica a ValorMercadoriaMoeda
           lErro = Valor_Positivo_Critica(ValorMercadoriaMoeda.Text)
           If lErro <> SUCESSO Then gError 196609
           
            If StrParaInt(MoedaMercadoria.Text) = 1 Then
                dTaxa = StrParaDbl(TaxaMoeda1.Text)
            Else
                dTaxa = StrParaDbl(TaxaMoeda2.Text)
            End If
           
            If dTaxa <> 0 Then
                ValorMercadoriaEmReal.Text = Format(StrParaDbl(ValorMercadoriaMoeda.Text) * dTaxa, "STANDARD")
                Call ValorMercadoriaEmReal_Validate(bSGECancelDummy)
            End If
    
        End If

        iMercEmMoedaAlterada = 0

    End If
    
    Exit Sub

Erro_ValorMercadoriaMoeda_Validate:

    Cancel = True

    Select Case gErr

        Case 196609

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196610)

    End Select

    Exit Sub

End Sub

Private Sub ValorMercadoriaMoeda_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ValorMercadoriaMoeda, iAlterado)
    
End Sub

Private Sub ValorMercadoriaMoeda_Change()
    iAlterado = REGISTRO_ALTERADO
    iMercEmMoedaAlterada = REGISTRO_ALTERADO
End Sub

Private Sub ValorFreteInternacMoeda_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dTaxa As Double

On Error GoTo Erro_ValorFreteInternacMoeda_Validate

    If iFreteEmMoedaAlterada <> 0 Or iMoedaFreteAnt <> StrParaInt(MoedaFrete.Text) Then
    
        iMoedaFreteAnt = StrParaInt(MoedaFrete.Text)
    
        'Verifica se ValorFreteInternacMoeda está preenchida
        If Len(Trim(ValorFreteInternacMoeda.Text)) <> 0 Then
    
            'Critica a ValorFreteInternacMoeda
            lErro = Valor_Positivo_Critica(ValorFreteInternacMoeda.Text)
            If lErro <> SUCESSO Then gError 196611
           
        End If
        
        If StrParaInt(MoedaFrete.Text) = 1 Then
            dTaxa = StrParaDbl(TaxaMoeda1.Text)
        Else
            dTaxa = StrParaDbl(TaxaMoeda2.Text)
        End If

        If dTaxa <> 0 Then
            ValorFreteInternacEmReal.Text = Format(StrParaDbl(ValorFreteInternacMoeda.Text) * dTaxa, "STANDARD")
            Call ValorFreteInternacEmReal_Validate(bSGECancelDummy)
        End If
    
        iFreteEmMoedaAlterada = 0
        
    End If
    
    Exit Sub

Erro_ValorFreteInternacMoeda_Validate:

    Cancel = True

    Select Case gErr

        Case 196611

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196612)

    End Select

    Exit Sub

End Sub

Private Sub ValorFreteInternacMoeda_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ValorFreteInternacMoeda, iAlterado)
    
End Sub

Private Sub ValorFreteInternacMoeda_Change()
    iAlterado = REGISTRO_ALTERADO
    iFreteEmMoedaAlterada = REGISTRO_ALTERADO
End Sub

Private Sub ValorSeguroInternacMoeda_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dTaxa As Double

On Error GoTo Erro_ValorSeguroInternacMoeda_Validate

    If iSeguroEmMoedaAlterada <> 0 Or iMoedaSeguroAnt <> StrParaInt(MoedaSeguro.Text) Then
    
        iMoedaSeguroAnt = StrParaInt(MoedaSeguro.Text)

        'Verifica se ValorSeguroInternacMoeda está preenchida
        If Len(Trim(ValorSeguroInternacMoeda.Text)) <> 0 Then
    
           'Critica a ValorSeguroInternacMoeda
           lErro = Valor_Positivo_Critica(ValorSeguroInternacMoeda.Text)
           If lErro <> SUCESSO Then gError 196613
           
        End If
        
        If StrParaInt(MoedaSeguro.Text) = 1 Then
            dTaxa = StrParaDbl(TaxaMoeda1.Text)
        Else
            dTaxa = StrParaDbl(TaxaMoeda2.Text)
        End If
        
        If dTaxa <> 0 Then
            ValorSeguroInternacEmReal.Text = Format(StrParaDbl(ValorSeguroInternacMoeda.Text) * dTaxa, "STANDARD")
            Call ValorSeguroInternacEmReal_Validate(bSGECancelDummy)
        End If
    
        iSeguroEmMoedaAlterada = 0
        
    End If

    Exit Sub

Erro_ValorSeguroInternacMoeda_Validate:

    Cancel = True

    Select Case gErr

        Case 196613

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196614)

    End Select

    Exit Sub

End Sub

Private Sub ValorSeguroInternacMoeda_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ValorSeguroInternacMoeda, iAlterado)
    
End Sub

Private Sub ValorSeguroInternacMoeda_Change()
    iAlterado = REGISTRO_ALTERADO
    iSeguroEmMoedaAlterada = REGISTRO_ALTERADO
End Sub

Private Sub ValorMercadoriaEmReal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorMercadoriaEmReal_Validate

    If iMercEmRealAlterada <> 0 Then
    
        'Verifica se ValorMercadoriaEmReal está preenchida
        If Len(Trim(ValorMercadoriaEmReal.Text)) <> 0 Then
    
           'Critica a ValorMercadoriaEmReal
           lErro = Valor_Positivo_Critica(ValorMercadoriaEmReal.Text)
           If lErro <> SUCESSO Then gError 196615
           
        End If
        
        Call RecalcularCIFs
        
        iMercEmRealAlterada = 0

    End If
    
    Exit Sub

Erro_ValorMercadoriaEmReal_Validate:

    Cancel = True

    Select Case gErr

        Case 196615

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196616)

    End Select

    Exit Sub

End Sub

Private Sub ValorMercadoriaEmReal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ValorMercadoriaEmReal, iAlterado)
    
End Sub

Private Sub ValorMercadoriaEmReal_Change()
    iAlterado = REGISTRO_ALTERADO
    iMercEmRealAlterada = REGISTRO_ALTERADO
End Sub

Private Sub ValorFreteInternacEmReal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorFreteInternacEmReal_Validate

    If iFreteEmRealAlterada <> 0 Then
    
        'Verifica se ValorFreteInternacEmReal está preenchida
        If Len(Trim(ValorFreteInternacEmReal.Text)) <> 0 Then
    
           'Critica a ValorFreteInternacEmReal
           lErro = Valor_NaoNegativo_Critica(ValorFreteInternacEmReal.Text)
           If lErro <> SUCESSO Then gError 196617
           
        End If

        Call RecalcularCIFs
        
        iFreteEmRealAlterada = 0
        
    End If
    
    Exit Sub

Erro_ValorFreteInternacEmReal_Validate:

    Cancel = True

    Select Case gErr

        Case 196617

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196618)

    End Select

    Exit Sub

End Sub

Private Sub ValorFreteInternacEmReal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ValorFreteInternacEmReal, iAlterado)
    
End Sub

Private Sub ValorFreteInternacEmReal_Change()
    iAlterado = REGISTRO_ALTERADO
    iFreteEmRealAlterada = REGISTRO_ALTERADO
End Sub

Private Sub ValorSeguroInternacEmReal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorSeguroInternacEmReal_Validate

    If iSeguroEmRealAlterada <> 0 Then
    
        'Verifica se ValorSeguroInternacEmReal está preenchida
        If Len(Trim(ValorSeguroInternacEmReal.Text)) <> 0 Then
    
           'Critica a ValorSeguroInternacEmReal
           lErro = Valor_NaoNegativo_Critica(ValorSeguroInternacEmReal.Text)
           If lErro <> SUCESSO Then gError 196619
           
        End If

        Call RecalcularCIFs
        
        iSeguroEmRealAlterada = 0
        
    End If
    
    Exit Sub

Erro_ValorSeguroInternacEmReal_Validate:

    Cancel = True

    Select Case gErr

        Case 196619

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196620)

    End Select

    Exit Sub

End Sub

Private Sub ValorSeguroInternacEmReal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ValorSeguroInternacEmReal, iAlterado)
    
End Sub

Private Sub ValorSeguroInternacEmReal_Change()
    iAlterado = REGISTRO_ALTERADO
    iSeguroEmRealAlterada = REGISTRO_ALTERADO
End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objDIInfo As ClassDIInfo

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objDIInfo = obj1

    'Mostra os dados do DIInfo na tela
    lErro = Traz_DIInfo_Tela(objDIInfo)
    If lErro <> SUCESSO Then gError 196621

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case 196621

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196622)

    End Select

    Exit Sub

End Sub

Private Sub LabelNumero_Click()

Dim lErro As Long
Dim objDIInfo As New ClassDIInfo
Dim colSelecao As New Collection

On Error GoTo Erro_LabelNumero_Click

    'Verifica se o Numero foi preenchido
    If Len(Trim(Numero.Text)) <> 0 Then

        objDIInfo.sNumero = Numero.Text

    End If

    Call Chama_Tela("DIInfoLista", colSelecao, objDIInfo, objEventoNumero)

    Exit Sub

Erro_LabelNumero_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196623)

    End Select

    Exit Sub

End Sub

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        FrameOpcao(Opcao.SelectedItem.Index).Visible = True
        FrameOpcao(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

    End If

End Sub

Private Function Inicializa_GridAdicao(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Clas. Fiscal")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("II %")
    objGrid.colColuna.Add ("IPI %")
    objGrid.colColuna.Add ("ICMS %")
    objGrid.colColuna.Add ("ICMS % Red.")
    objGrid.colColuna.Add ("PIS %")
    objGrid.colColuna.Add ("COFINS %")
    objGrid.colColuna.Add ("Valor Aduaneiro")
    objGrid.colColuna.Add ("II R$")
    objGrid.colColuna.Add ("Base IPI")
    objGrid.colColuna.Add ("IPI R$")
    objGrid.colColuna.Add ("D.Aduaneira")
    objGrid.colColuna.Add ("Siscomex")
    objGrid.colColuna.Add ("Base PIS")
    objGrid.colColuna.Add ("PIS R$")
    objGrid.colColuna.Add ("Base COFINS")
    objGrid.colColuna.Add ("COFINS R$")
    objGrid.colColuna.Add ("Base ICMS")
    objGrid.colColuna.Add ("ICMS R$")
    objGrid.colColuna.Add ("Cód. Fabricante")
    objGrid.colColuna.Add ("N.Ato Drawback")

    'Controles que participam do Grid
    objGrid.colCampo.Add (IPICodigo.Name)
    objGrid.colCampo.Add (IPIDescricao.Name)
    objGrid.colCampo.Add (IIAliquota.Name)
    objGrid.colCampo.Add (IPIAliquota.Name)
    objGrid.colCampo.Add (ICMSAliquota.Name)
    objGrid.colCampo.Add (ICMSPercRedBase.Name)
    objGrid.colCampo.Add (PISAliquota.Name)
    objGrid.colCampo.Add (COFINSAliquota.Name)
    objGrid.colCampo.Add (AdicaoValorAduaneiro.Name)
    objGrid.colCampo.Add (IIValor.Name)
    objGrid.colCampo.Add (IPIBase.Name)
    objGrid.colCampo.Add (IPIValor.Name)
    objGrid.colCampo.Add (DespAdua.Name)
    objGrid.colCampo.Add (TaxaSiscomex.Name)
    objGrid.colCampo.Add (PISBase.Name)
    objGrid.colCampo.Add (PISValor.Name)
    objGrid.colCampo.Add (COFINSBase.Name)
    objGrid.colCampo.Add (COFINSValor.Name)
    objGrid.colCampo.Add (ICMSBase.Name)
    objGrid.colCampo.Add (ICMSValor.Name)
    objGrid.colCampo.Add (CodFabricante.Name)
    objGrid.colCampo.Add (NumDrawback.Name)

    'Colunas do Grid
    iGrid_IPICodigo_Col = 1
    iGrid_IPIDescricao_Col = 2
    iGrid_IIAliquota_Col = 3
    iGrid_IPIAliquota_Col = 4
    iGrid_ICMSAliquota_Col = 5
    iGrid_ICMSPercRedBase_Col = 6
    iGrid_PISAliquota_Col = 7
    iGrid_COFINSAliquota_Col = 8
    iGrid_AdicaoValorAduaneiro_Col = 9
    iGrid_IIValor_Col = 10
    iGrid_IPIBase_Col = 11
    iGrid_IPIValor_Col = 12
    iGrid_DespAdua_Col = 13
    iGrid_TaxaSiscomex_Col = 14
    iGrid_PISBase_Col = 15
    iGrid_PISValor_Col = 16
    iGrid_COFINSBase_Col = 17
    iGrid_COFINSValor_Col = 18
    iGrid_ICMSBase_Col = 19
    iGrid_ICMSValor_Col = 20
    iGrid_CodFabricante_Col = 21
    iGrid_NumDrawback_Col = 22

    objGrid.objGrid = GridAdicao

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_LINHAS_GRID_ADICAO + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridAdicao.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridAdicao = SUCESSO

End Function

Private Function Inicializa_GridItens(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Adição")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Descricao")
    objGrid.colColuna.Add ("UM")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Unit. FOB")
    objGrid.colColuna.Add ("Unit. FOB R$")
    objGrid.colColuna.Add ("Unit. CIF")
    objGrid.colColuna.Add ("Unit. CIF R$")
    objGrid.colColuna.Add ("Total FOB")
    objGrid.colColuna.Add ("Total FOB R$")
    objGrid.colColuna.Add ("Total CIF")
    objGrid.colColuna.Add ("Total CIF R$")
    objGrid.colColuna.Add ("Manual")
    objGrid.colColuna.Add ("Peso Bruto")
    objGrid.colColuna.Add ("Peso Líquido")
    objGrid.colColuna.Add ("IPI Unit.")

    'Controles que participam do Grid
    objGrid.colCampo.Add (AdicaoItem.Name)
    objGrid.colCampo.Add (Produto.Name)
    objGrid.colCampo.Add (Descricao.Name)
    objGrid.colCampo.Add (UnidadeMed.Name)
    objGrid.colCampo.Add (Quantidade.Name)
    objGrid.colCampo.Add (ValorUnitFOBNaMoeda.Name)
    objGrid.colCampo.Add (ValorUnitFOBEmReal.Name)
    objGrid.colCampo.Add (ValorUnitCIFNaMoeda.Name)
    objGrid.colCampo.Add (ValorUnitCIFEmReal.Name)
    objGrid.colCampo.Add (ValorTotalFOBNaMoeda.Name)
    objGrid.colCampo.Add (ValorTotalFOBEmReal.Name)
    objGrid.colCampo.Add (ValorTotalCIFNaMoeda.Name)
    objGrid.colCampo.Add (ValorTotalCIFEmReal.Name)
    objGrid.colCampo.Add (TotalCIFEmRealManual.Name)
    objGrid.colCampo.Add (ItemPesoBruto.Name)
    objGrid.colCampo.Add (ItemPesoLiq.Name)
    objGrid.colCampo.Add (IPIValorUnitario.Name)

    'Colunas do Grid
    iGrid_AdicaoItem_Col = 1
    iGrid_Produto_Col = 2
    iGrid_Descricao_Col = 3
    iGrid_UnidadeMed_Col = 4
    iGrid_Quantidade_Col = 5
    iGrid_ValorUnitFOBNaMoeda_Col = 6
    iGrid_ValorUnitFOBEmReal_Col = 7
    iGrid_ValorUnitCIFNaMoeda_Col = 8
    iGrid_ValorUnitCIFEmReal_Col = 9
    iGrid_ValorTotalFOBNaMoeda_Col = 10
    iGrid_ValorTotalFOBEmReal_Col = 11
    iGrid_ValorTotalCIFNaMoeda_Col = 12
    iGrid_ValorTotalCIFEmReal_Col = 13
    iGrid_TotalCIFEmRealManual_Col = 14
    iGrid_ItemPesoBruto_Col = 15
    iGrid_ItemPesoLiq_Col = 16
    iGrid_IPIValorUnitario_Col = 17

    objGrid.objGrid = GridItens

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_LINHAS_GRID_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridItens = SUCESSO

End Function

Private Sub GridAdicao_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridAdicao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAdicao, iAlterado)
    End If

End Sub

Private Sub GridAdicao_GotFocus()
    Call Grid_Recebe_Foco(objGridAdicao)
End Sub

Private Sub GridAdicao_EnterCell()
    Call Grid_Entrada_Celula(objGridAdicao, iAlterado)
End Sub

Private Sub GridAdicao_LeaveCell()
    Call Saida_Celula(objGridAdicao)
End Sub

Private Sub GridAdicao_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridAdicao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAdicao, iAlterado)
    End If

End Sub

Private Sub GridAdicao_RowColChange()
    Call Grid_RowColChange(objGridAdicao)
End Sub

Private Sub GridAdicao_Scroll()
    Call Grid_Scroll(objGridAdicao)
End Sub

Private Sub GridAdicao_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer
Dim iIndice As Integer
Dim iLinha As Integer
Dim iItem As Integer

On Error GoTo Erro_GridAdicao_KeyDown

    iLinhasExistentesAnterior = objGridAdicao.iLinhasExistentes
    iItemAtual = GridAdicao.Row

    Call Grid_Trata_Tecla1(KeyCode, objGridAdicao)

    'Se exclui uma linha de itens
    If objGridAdicao.iLinhasExistentes < iLinhasExistentesAnterior Then
        
        For iLinha = objGridItens.iLinhasExistentes To 1 Step -1
            iItem = StrParaInt(GridItens.TextMatrix(iLinha, iGrid_AdicaoItem_Col))
            
            'Se a adição foi excluida
            If iItem = iItemAtual Then
            
                Call Grid_Exclui_Linha(objGridItens, iItem)
                
            ElseIf iItem > iItemAtual Then
            
                GridItens.TextMatrix(iLinha, iGrid_AdicaoItem_Col) = CStr(iItem - 1)
                
            End If
    
        Next
        
        Call Calcula_Valores_Adicao(0, 0)
    
    End If
    
    Exit Sub
    
Erro_GridAdicao_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196624)

    End Select

    Exit Sub

End Sub

Private Sub GridAdicao_LostFocus()
    Call Grid_Libera_Foco(objGridAdicao)
End Sub

Private Sub IPICodigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub IPICodigo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub IPICodigo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub IPICodigo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = IPICodigo
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub IPIAliquota_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub IPIAliquota_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub IPIAliquota_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub IPIAliquota_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = IPIAliquota
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ICMSAliquota_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ICMSAliquota_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub ICMSAliquota_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub ICMSAliquota_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = ICMSAliquota
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub IIAliquota_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub IIAliquota_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub IIAliquota_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub IIAliquota_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = IIAliquota
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

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
    If GridItens.Row <> 0 Then
        Call Atualiza_Totais_Adicao(GridItens.Row)
        Call Exibe_CampoDet_Grid(objGridItens, iGrid_Descricao_Col, DescDet)
    End If
End Sub

Private Sub GridItens_Scroll()
    Call Grid_Scroll(objGridItens)
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer

On Error GoTo Erro_GridItens_KeyDown

    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes
    iItemAtual = GridItens.Row

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

    'Se exclui uma linha de itens
    If objGridItens.iLinhasExistentes < iLinhasExistentesAnterior Then
        
        Call Calcula_Valores_Adicao(0, 0)
    
    End If
    
    Exit Sub
    
Erro_GridItens_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196625)

    End Select

    Exit Sub
    
End Sub

Private Sub GridItens_LostFocus()
    Call Grid_Libera_Foco(objGridItens)
End Sub

Private Sub Produto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Produto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Descricao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Descricao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Descricao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Descricao
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Quantidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Quantidade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UnidadeMed_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UnidadeMed_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub UnidadeMed_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub UnidadeMed_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UnidadeMed
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorUnitFOBNaMoeda_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorUnitFOBNaMoeda_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub ValorUnitFOBNaMoeda_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub ValorUnitFOBNaMoeda_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ValorUnitFOBNaMoeda
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorUnitFOBEmReal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorUnitFOBEmReal_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub ValorUnitFOBEmReal_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub ValorUnitFOBEmReal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ValorUnitFOBEmReal
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorUnitCIFNaMoeda_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorUnitCIFNaMoeda_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub ValorUnitCIFNaMoeda_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub ValorUnitCIFNaMoeda_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ValorUnitCIFNaMoeda
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorUnitCIFEmReal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorUnitCIFEmReal_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub ValorUnitCIFEmReal_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub ValorUnitCIFEmReal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ValorUnitCIFEmReal
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub IPIValorUnitario_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub IPIValorUnitario_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub IPIValorUnitario_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub IPIValorUnitario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = IPIValorUnitario
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorTotalFOBNaMoeda_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorTotalFOBNaMoeda_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub ValorTotalFOBNaMoeda_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub ValorTotalFOBNaMoeda_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ValorTotalFOBNaMoeda
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorTotalFOBEmReal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorTotalFOBEmReal_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub ValorTotalFOBEmReal_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub ValorTotalFOBEmReal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ValorTotalFOBEmReal
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorTotalCIFNaMoeda_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorTotalCIFNaMoeda_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub ValorTotalCIFNaMoeda_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub ValorTotalCIFNaMoeda_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ValorTotalCIFNaMoeda
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorTotalCIFEmReal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorTotalCIFEmReal_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub ValorTotalCIFEmReal_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub ValorTotalCIFEmReal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ValorTotalCIFEmReal
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TotalCIFEmRealManual_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TotalCIFEmRealManual_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub TotalCIFEmRealManual_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub TotalCIFEmRealManual_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = TotalCIFEmRealManual
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'GridAdicao
        If objGridInt.objGrid.Name = GridAdicao.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_TaxaSiscomex_Col, iGrid_DespAdua_Col, iGrid_AdicaoValorAduaneiro_Col, iGrid_IIValor_Col, iGrid_IPIBase_Col, iGrid_IPIValor_Col, iGrid_PISBase_Col, iGrid_PISValor_Col, iGrid_ICMSValor_Col, iGrid_ICMSBase_Col, iGrid_COFINSValor_Col, iGrid_PISValor_Col, iGrid_COFINSBase_Col

                    lErro = Saida_Celula_Valor(objGridInt, Me.ActiveControl)
                    If lErro <> SUCESSO Then gError 196626
                
                Case iGrid_PISAliquota_Col, iGrid_COFINSAliquota_Col, iGrid_IPIAliquota_Col, iGrid_IIAliquota_Col, iGrid_ICMSAliquota_Col, iGrid_ICMSPercRedBase_Col

                    lErro = Saida_Celula_Percentual(objGridInt, Me.ActiveControl)
                    If lErro <> SUCESSO Then gError 196627
                
                Case iGrid_IPICodigo_Col

                    lErro = Saida_Celula_IPICodigo(objGridInt)
                    If lErro <> SUCESSO Then gError 196628
                    
                Case iGrid_CodFabricante_Col, iGrid_NumDrawback_Col

                    lErro = Saida_Celula_Padrao(objGridInt, Me.ActiveControl)
                    If lErro <> SUCESSO Then gError 196630

            End Select
                    
        End If

        'GridItens
        If objGridInt.objGrid.Name = GridItens.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_Produto_Col

                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 196629

                Case iGrid_Descricao_Col, iGrid_UnidadeMed_Col, iGrid_AdicaoItem_Col

                    lErro = Saida_Celula_Padrao(objGridInt, Me.ActiveControl)
                    If lErro <> SUCESSO Then gError 196630

                Case iGrid_Quantidade_Col, iGrid_ItemPesoBruto_Col, iGrid_ItemPesoLiq_Col

                    lErro = Saida_Celula_Quantidade(objGridInt, Me.ActiveControl)
                    If lErro <> SUCESSO Then gError 196631

                Case iGrid_ValorUnitFOBNaMoeda_Col, iGrid_ValorUnitFOBEmReal_Col, iGrid_ValorUnitCIFNaMoeda_Col, iGrid_ValorUnitCIFEmReal_Col, iGrid_ValorTotalFOBEmReal_Col, iGrid_ValorTotalFOBNaMoeda_Col, iGrid_ValorTotalCIFNaMoeda_Col, iGrid_ValorTotalCIFEmReal_Col, iGrid_IPIValorUnitario_Col

                    lErro = Saida_Celula_Valor(objGridInt, Me.ActiveControl)
                    If lErro <> SUCESSO Then gError 196632

            End Select
                    
        End If

        'GridDespesas
        If objGridInt.objGrid.Name = GridDespesas.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_ComplTipo_Col

                    lErro = Saida_Celula_ComplTipo(objGridInt)
                    If lErro <> SUCESSO Then gError 196634
                
                Case iGrid_ComplDescricao_Col

                    lErro = Saida_Celula_Padrao(objGridInt, ComplDescricao)
                    If lErro <> SUCESSO Then gError 196635
                
                Case iGrid_ComplValor_Col

                    lErro = Saida_Celula_Valor(objGridInt, ComplValor)
                    If lErro <> SUCESSO Then gError 196636
                    
                Case iGrid_ComplPerc_Col
                
                    lErro = Saida_Celula_Percentual(objGridInt, Me.ActiveControl)
                    If lErro <> SUCESSO Then gError 196636

                Case iGrid_ComplDias_Col
                
                    lErro = Saida_Celula_Inteiro(objGridInt, Me.ActiveControl)
                    If lErro <> SUCESSO Then gError 196636
                
            End Select
                    
        End If
        
        'GridItensPC
        If objGridInt.objGrid.Name = GridItensPC.Name Then

            lErro = Saida_Celula_GridItensPC(objGridInt)
            If lErro <> SUCESSO Then gError 210567

        End If
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 196637

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 196626 To 196636, 210567

        Case 196637
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196638)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatadoBenef As String
Dim iProdutoPreenchidoBenef As Integer
Dim objProduto As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String
Dim sUnidadeMed As String
Dim iIndice As Integer
Dim objTipoDespesa As New ClassTipoImportCompl

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa o controle da coluna em questão
    Select Case objControl.Name
    
        'Produto
        Case Produto.Name
            'If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Produto_Col))) > 0 Then
                'Produto.Enabled = False
            'Else
                Produto.Enabled = True
            'End If
    
        'Unidade de Medida
        Case UnidadeMed.Name

            UnidadeMed.Clear

            'Guarda a UM que está no Grid
            sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)

            lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 196639
           
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                UnidadeMed.Enabled = False
            Else
                UnidadeMed.Enabled = True

                objProduto.sCodigo = sProdutoFormatado
                
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 196640

                objClasseUM.iClasse = objProduto.iClasseUM
                
                'Lâ as Unidades de Medidas da Classe do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 196641
                
                'Carrega a combo de UM
                For Each objUM In colSiglas
                    UnidadeMed.AddItem objUM.sSigla
                Next
                
                'Seleciona na UM que está preenchida
                UnidadeMed.Text = sUM
                
                If Len(Trim(sUM)) > 0 Then
                    lErro = Combo_Item_Igual(UnidadeMed)
                    If lErro <> SUCESSO And lErro <> 12253 Then gError 196642
                End If
            
            End If
            
        Case AdicaoItem.Name, CodigoPC.Name
            objControl.Enabled = True
            
            
        Case IPIDescricao.Name, Descricao.Name, ValorUnitFOBEmReal.Name, ValorTotalCIFNaMoeda.Name, _
                ValorUnitCIFEmReal.Name, ValorUnitCIFNaMoeda.Name, ValorTotalFOBNaMoeda.Name, DescProdutoPC.Name, _
                UMPC.Name, DataPC.Name
                objControl.Enabled = False

        Case Quantidade.Name, ValorUnitFOBNaMoeda.Name, _
                ValorTotalFOBEmReal.Name, ValorTotalCIFEmReal.Name, TotalCIFEmRealManual.Name, ItemPesoBruto.Name, ItemPesoLiq.Name, IPIValorUnitario.Name

            If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Produto_Col))) > 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
        Case IPICodigo.Name
        
            If Len(Trim(GridAdicao.TextMatrix(iLinha, iGrid_IPICodigo_Col))) > 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case AdicaoValorAduaneiro.Name, IIAliquota.Name, IIValor.Name, IPIBase.Name, IPIAliquota.Name, _
                IPIValor.Name, PISBase.Name, PISAliquota.Name, PISValor.Name, COFINSBase.Name, COFINSAliquota.Name, _
                COFINSValor.Name, ICMSBase.Name, ICMSAliquota.Name, ICMSValor.Name, DespAdua.Name, TaxaSiscomex.Name, CodFabricante.Name, ICMSPercRedBase.Name
        
            If Len(Trim(GridAdicao.TextMatrix(iLinha, iGrid_IPICodigo_Col))) > 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
        Case ComplTipo.Name
        
            If Len(Trim(GridDespesas.TextMatrix(iLinha, iGrid_ComplTipo_Col))) > 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case ComplDescricao.Name, ComplValor.Name, ComplPerc.Name, ComplDias.Name
        
            If Len(Trim(GridDespesas.TextMatrix(iLinha, iGrid_ComplTipo_Col))) > 0 Then
            
                objTipoDespesa.iCodigo = StrParaInt(GridDespesas.TextMatrix(iLinha, iGrid_ComplTipo_Col))
                
                lErro = CF("TiposImportCompl_Le", objTipoDespesa)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 196642
            
                If objControl.Name = ComplValor.Name Then
                    If objTipoDespesa.iAceitaValor = MARCADO Then
                        objControl.Enabled = True
                    Else
                        objControl.Enabled = False
                    End If
                ElseIf objControl.Name = ComplPerc.Name Then
                    If objTipoDespesa.iAceitaPerc = MARCADO Then
                        objControl.Enabled = True
                    Else
                        objControl.Enabled = False
                    End If
                ElseIf objControl.Name = ComplDias.Name Then
                    If objTipoDespesa.iAceitaDias = MARCADO Then
                        objControl.Enabled = True
                    Else
                        objControl.Enabled = False
                    End If
                Else
                    objControl.Enabled = True
                End If

            Else
                objControl.Enabled = False
            End If


        Case ProdutoPC.Name
            If Len(Trim(GridItensPC.TextMatrix(iLinha, iGrid_CodigoPC_Col))) > 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If

        Case QuantPC.Name

            If Len(Trim(GridItensPC.TextMatrix(iLinha, iGrid_ProdutoPC_Col))) > 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If


        Case Else
            objControl.Enabled = True

    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case 196639 To 196642

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196643)

    End Select

    Exit Sub

End Sub

Function Preenche_GridAdicao_Tela(objDIInfo As ClassDIInfo) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objAdicaoDI As ClassAdicaoDI
Dim objIPICodigo As ClassClassificacaoFiscal

On Error GoTo Erro_Preenche_GridAdicao_Tela

    Call Grid_Limpa(objGridAdicao)

    iLinha = 0
    
    For Each objAdicaoDI In objDIInfo.colAdicoesDI
    
        iLinha = iLinha + 1
    
        GridAdicao.TextMatrix(iLinha, iGrid_COFINSAliquota_Col) = Format(objAdicaoDI.dCOFINSAliquota, "Percent")
        GridAdicao.TextMatrix(iLinha, iGrid_COFINSBase_Col) = Format(objAdicaoDI.dCOFINSBase, "STANDARD")
        GridAdicao.TextMatrix(iLinha, iGrid_COFINSValor_Col) = Format(objAdicaoDI.dCOFINSValor, "STANDARD")
        GridAdicao.TextMatrix(iLinha, iGrid_ICMSAliquota_Col) = Format(objAdicaoDI.dICMSAliquota, "Percent")
        GridAdicao.TextMatrix(iLinha, iGrid_ICMSPercRedBase_Col) = Format(objAdicaoDI.dICMSPercRedBase, "Percent")
        GridAdicao.TextMatrix(iLinha, iGrid_ICMSBase_Col) = Format(objAdicaoDI.dICMSBase, "STANDARD")
        GridAdicao.TextMatrix(iLinha, iGrid_ICMSValor_Col) = Format(objAdicaoDI.dICMSValor, "STANDARD")
        GridAdicao.TextMatrix(iLinha, iGrid_IIAliquota_Col) = Format(objAdicaoDI.dIIAliquota, "Percent")
        GridAdicao.TextMatrix(iLinha, iGrid_IIValor_Col) = Format(objAdicaoDI.dIIValor, "STANDARD")
        GridAdicao.TextMatrix(iLinha, iGrid_IPIAliquota_Col) = Format(objAdicaoDI.dIPIAliquota, "Percent")
        GridAdicao.TextMatrix(iLinha, iGrid_IPIBase_Col) = Format(objAdicaoDI.dIPIBase, "STANDARD")
        GridAdicao.TextMatrix(iLinha, iGrid_IPIValor_Col) = Format(objAdicaoDI.dIPIValor, "STANDARD")
        GridAdicao.TextMatrix(iLinha, iGrid_PISAliquota_Col) = Format(objAdicaoDI.dPISAliquota, "Percent")
        GridAdicao.TextMatrix(iLinha, iGrid_PISBase_Col) = Format(objAdicaoDI.dPISBase, "STANDARD")
        GridAdicao.TextMatrix(iLinha, iGrid_PISValor_Col) = Format(objAdicaoDI.dPISValor, "STANDARD")
        GridAdicao.TextMatrix(iLinha, iGrid_AdicaoValorAduaneiro_Col) = Format(objAdicaoDI.dValorAduaneiro, "STANDARD")
        GridAdicao.TextMatrix(iLinha, iGrid_IPICodigo_Col) = objAdicaoDI.sIPICodigo
        
        GridAdicao.TextMatrix(iLinha, iGrid_DespAdua_Col) = Format(objAdicaoDI.dDespesaAduaneira, "STANDARD")
        GridAdicao.TextMatrix(iLinha, iGrid_TaxaSiscomex_Col) = Format(objAdicaoDI.dTaxaSiscomex, "STANDARD")
        GridAdicao.TextMatrix(iLinha, iGrid_CodFabricante_Col) = objAdicaoDI.sCodFabricante
        GridAdicao.TextMatrix(iLinha, iGrid_NumDrawback_Col) = objAdicaoDI.sNumDrawback
        
        Set objIPICodigo = New ClassClassificacaoFiscal
        
        objIPICodigo.sCodigo = objAdicaoDI.sIPICodigo
        
        lErro = CF("ClassificacaoFiscal_Le", objIPICodigo)
        If lErro <> SUCESSO And lErro <> 123494 Then gError 196644

        GridAdicao.TextMatrix(iLinha, iGrid_IPIDescricao_Col) = objIPICodigo.sDescricao

    Next

    objGridAdicao.iLinhasExistentes = objDIInfo.colAdicoesDI.Count
    
    Call Carrega_AdicaoItem
    
    lErro = Preenche_GridItens_Tela(objDIInfo)
    If lErro <> SUCESSO Then gError 196645
        
    Preenche_GridAdicao_Tela = SUCESSO

    Exit Function

Erro_Preenche_GridAdicao_Tela:

    Preenche_GridAdicao_Tela = gErr

    Select Case gErr
    
        Case 196644, 196645

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196646)

    End Select

    Exit Function

End Function

Function Move_GridAdicao_Memoria(objDIInfo As ClassDIInfo) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objAdicaoDI As ClassAdicaoDI

On Error GoTo Erro_Move_GridAdicao_Memoria

    For iLinha = 1 To objGridAdicao.iLinhasExistentes
    
        Set objAdicaoDI = New ClassAdicaoDI
        
        objAdicaoDI.dCOFINSAliquota = PercentParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_COFINSAliquota_Col))
        objAdicaoDI.dCOFINSBase = StrParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_COFINSBase_Col))
        objAdicaoDI.dCOFINSValor = StrParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_COFINSValor_Col))
        objAdicaoDI.dICMSAliquota = PercentParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_ICMSAliquota_Col))
        objAdicaoDI.dICMSPercRedBase = PercentParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_ICMSPercRedBase_Col))
        objAdicaoDI.dICMSBase = StrParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_ICMSBase_Col))
        objAdicaoDI.dICMSValor = StrParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_ICMSValor_Col))
        objAdicaoDI.dIIAliquota = PercentParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_IIAliquota_Col))
        objAdicaoDI.dIIValor = StrParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_IIValor_Col))
        objAdicaoDI.dIPIAliquota = PercentParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_IPIAliquota_Col))
        objAdicaoDI.dIPIBase = StrParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_IPIBase_Col))
        objAdicaoDI.dIPIValor = StrParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_IPIValor_Col))
        objAdicaoDI.dPISAliquota = PercentParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_PISAliquota_Col))
        objAdicaoDI.dPISBase = StrParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_PISBase_Col))
        objAdicaoDI.dPISValor = StrParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_PISValor_Col))
        objAdicaoDI.dValorAduaneiro = StrParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_AdicaoValorAduaneiro_Col))
        objAdicaoDI.sIPICodigo = GridAdicao.TextMatrix(iLinha, iGrid_IPICodigo_Col)
        objAdicaoDI.iSeq = iLinha
        
        objAdicaoDI.dTaxaSiscomex = StrParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_TaxaSiscomex_Col))
        objAdicaoDI.dDespesaAduaneira = StrParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_DespAdua_Col))
        objAdicaoDI.sCodFabricante = GridAdicao.TextMatrix(iLinha, iGrid_CodFabricante_Col)
        objAdicaoDI.sNumDrawback = GridAdicao.TextMatrix(iLinha, iGrid_NumDrawback_Col)
        
        If Len(Trim(objAdicaoDI.sCodFabricante)) = 0 And Not gbTrazendoPC Then gError 206785
        
        objDIInfo.colAdicoesDI.Add objAdicaoDI
    
    Next
    
    lErro = Move_GridItens_Memoria(objDIInfo)
    If lErro <> SUCESSO Then gError 196647
    
    Move_GridAdicao_Memoria = SUCESSO

    Exit Function

Erro_Move_GridAdicao_Memoria:

    Move_GridAdicao_Memoria = gErr

    Select Case gErr
    
        Case 196647

        Case 206785
            Call Rotina_Erro(vbOKOnly, "ERRO_CODFABRICANTE_NAO_PREENCHIDO", gErr, iLinha)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196648)

    End Select

    Exit Function

End Function

Function Preenche_GridDespesas_Tela(objDIInfo As ClassDIInfo) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objDespesas As ClassImportCompl

On Error GoTo Erro_Preenche_GridDespesas_Tela

    Call Grid_Limpa(objGridDespesas)

    iLinha = 0
    
    For Each objDespesas In objDIInfo.colDespesasDI
    
        iLinha = iLinha + 1
    
        GridDespesas.TextMatrix(iLinha, iGrid_ComplTipo_Col) = objDespesas.iTipo
        GridDespesas.TextMatrix(iLinha, iGrid_ComplDescricao_Col) = objDespesas.sDescricao
        GridDespesas.TextMatrix(iLinha, iGrid_ComplValor_Col) = Format(objDespesas.dValor, "STANDARD")
        GridDespesas.TextMatrix(iLinha, iGrid_ComplPerc_Col) = Format(objDespesas.dPerc, "Percent")
        GridDespesas.TextMatrix(iLinha, iGrid_ComplDias_Col) = CStr(objDespesas.iDias)

    Next

    objGridDespesas.iLinhasExistentes = objDIInfo.colDespesasDI.Count
    
    Preenche_GridDespesas_Tela = SUCESSO

    Exit Function

Erro_Preenche_GridDespesas_Tela:

    Preenche_GridDespesas_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196649)

    End Select

    Exit Function

End Function

Function Move_GridDespesas_Memoria(objDIInfo As ClassDIInfo) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objDespesas As ClassImportCompl

On Error GoTo Erro_Move_GridDespesas_Memoria

    For iLinha = 1 To objGridDespesas.iLinhasExistentes
    
        Set objDespesas = New ClassImportCompl
        
        objDespesas.iTipoDocOrigem = IMPORTCOMPL_ORIGEM_DI

        objDespesas.iSeq = iLinha
        objDespesas.iTipo = StrParaInt(GridDespesas.TextMatrix(iLinha, iGrid_ComplTipo_Col))
        objDespesas.sDescricao = GridDespesas.TextMatrix(iLinha, iGrid_ComplDescricao_Col)
        objDespesas.dValor = StrParaDbl(GridDespesas.TextMatrix(iLinha, iGrid_ComplValor_Col))
        objDespesas.dPerc = PercentParaDbl(GridDespesas.TextMatrix(iLinha, iGrid_ComplPerc_Col))
        objDespesas.iDias = StrParaInt(GridDespesas.TextMatrix(iLinha, iGrid_ComplDias_Col))
        
        objDIInfo.colDespesasDI.Add objDespesas
    
    Next
    
    Move_GridDespesas_Memoria = SUCESSO

    Exit Function

Erro_Move_GridDespesas_Memoria:

    Move_GridDespesas_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196650)

    End Select

    Exit Function

End Function

Function Preenche_GridItens_Tela(objDIInfo As ClassDIInfo) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objItemAdicaoDI As ClassItemAdicaoDI
Dim objAdicaoDI As ClassAdicaoDI
Dim sProdutoEnxuto As String
Dim objProduto As ClassProduto

On Error GoTo Erro_Preenche_GridItens_Tela

    Call Grid_Limpa(objGridItens)

    iLinha = 0
    
    For Each objAdicaoDI In objDIInfo.colAdicoesDI
    
        For Each objItemAdicaoDI In objAdicaoDI.colItensAdicaoDI
    
            iLinha = iLinha + 1
            
            If Len(Trim(objItemAdicaoDI.sProduto)) > 0 Then
            
                lErro = Mascara_RetornaProdutoEnxuto(objItemAdicaoDI.sProduto, sProdutoEnxuto)
                If lErro <> SUCESSO Then gError 196651
    
                'Call StrParaMasked2(Produto, sProdutoEnxuto)
                Produto.PromptInclude = False
                Produto.Text = sProdutoEnxuto
                Produto.PromptInclude = True
                
                Set objProduto = New ClassProduto
        
                objProduto.sCodigo = objItemAdicaoDI.sProduto
                
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 196652
                
                GridItens.TextMatrix(iLinha, iGrid_Produto_Col) = Produto.Text
                
            End If
            
            GridItens.TextMatrix(iLinha, iGrid_Descricao_Col) = objItemAdicaoDI.sDescricao
            
            GridItens.TextMatrix(iLinha, iGrid_ItemPesoBruto_Col) = Formata_Estoque(objItemAdicaoDI.dPesoBruto)
            GridItens.TextMatrix(iLinha, iGrid_ItemPesoLiq_Col) = Formata_Estoque(objItemAdicaoDI.dPesoLiq)
            GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(objItemAdicaoDI.dQuantidade)
            GridItens.TextMatrix(iLinha, iGrid_ValorTotalCIFEmReal_Col) = Format(objItemAdicaoDI.dValorTotalCIFEmReal, "STANDARD")
            GridItens.TextMatrix(iLinha, iGrid_TotalCIFEmRealManual_Col) = CStr(objItemAdicaoDI.iTotalCIFEmRealManual)
            GridItens.TextMatrix(iLinha, iGrid_ValorTotalCIFNaMoeda_Col) = Format(objItemAdicaoDI.dValorTotalCIFNaMoeda, "STANDARD")
            GridItens.TextMatrix(iLinha, iGrid_ValorTotalFOBEmReal_Col) = Format(objItemAdicaoDI.dValorTotalFOBEmReal, "STANDARD")
            GridItens.TextMatrix(iLinha, iGrid_ValorTotalFOBNaMoeda_Col) = Format(objItemAdicaoDI.dValorTotalFOBNaMoeda, "#,##0.00#####")
            GridItens.TextMatrix(iLinha, iGrid_ValorUnitCIFEmReal_Col) = Format(objItemAdicaoDI.dValorUnitCIFEmReal, "STANDARD")
            GridItens.TextMatrix(iLinha, iGrid_ValorUnitCIFNaMoeda_Col) = Format(objItemAdicaoDI.dValorUnitCIFNaMoeda, "#,##0.00#####")
            GridItens.TextMatrix(iLinha, iGrid_ValorUnitFOBEmReal_Col) = Format(objItemAdicaoDI.dValorUnitFOBEmReal, "STANDARD")
            GridItens.TextMatrix(iLinha, iGrid_ValorUnitFOBNaMoeda_Col) = Format(objItemAdicaoDI.dValorUnitFOBNaMoeda, "#,##0.00#####")
            GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col) = objItemAdicaoDI.sUM
            GridItens.TextMatrix(iLinha, iGrid_AdicaoItem_Col) = CStr(objAdicaoDI.iSeq)
            GridItens.TextMatrix(iLinha, iGrid_IPIValorUnitario_Col) = Format(objItemAdicaoDI.dIPIUnidadePadraoValor, "STANDARD")

        Next

    Next

    objGridItens.iLinhasExistentes = iLinha
    Call Grid_Refresh_Checkbox(objGridItens)
    
    Call Atualiza_Totais_Itens
    
    Preenche_GridItens_Tela = SUCESSO

    Exit Function

Erro_Preenche_GridItens_Tela:

    Preenche_GridItens_Tela = gErr

    Select Case gErr
    
        Case 196651, 196652

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196653)

    End Select

    Exit Function

End Function

Function Move_GridItens_Memoria(objDIInfo As ClassDIInfo) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objAdicaoDI As ClassAdicaoDI
Dim objItemAdicaoDI As ClassItemAdicaoDI
Dim bAchou As Boolean
Dim iItem As Integer
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Move_GridItens_Memoria

    For iLinha = 1 To objGridItens.iLinhasExistentes
    
        Set objItemAdicaoDI = New ClassItemAdicaoDI
    
        iItem = StrParaInt(GridItens.TextMatrix(iLinha, iGrid_AdicaoItem_Col))
        
        bAchou = False
        For Each objAdicaoDI In objDIInfo.colAdicoesDI
            If objAdicaoDI.iSeq = iItem Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then gError 196654
    
        'Formata o produto
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 196655

        objItemAdicaoDI.sProduto = sProduto
        
        objItemAdicaoDI.dPesoBruto = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ItemPesoBruto_Col))
        objItemAdicaoDI.dPesoLiq = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ItemPesoLiq_Col))
        objItemAdicaoDI.dQuantidade = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))
        objItemAdicaoDI.dValorTotalCIFEmReal = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorTotalCIFEmReal_Col))
        objItemAdicaoDI.iTotalCIFEmRealManual = StrParaInt(GridItens.TextMatrix(iLinha, iGrid_TotalCIFEmRealManual_Col))
        objItemAdicaoDI.dValorTotalCIFNaMoeda = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorTotalCIFNaMoeda_Col))
        objItemAdicaoDI.dValorTotalFOBEmReal = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorTotalFOBEmReal_Col))
        objItemAdicaoDI.dValorTotalFOBNaMoeda = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorTotalFOBNaMoeda_Col))
        objItemAdicaoDI.dValorUnitCIFEmReal = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitCIFEmReal_Col))
        objItemAdicaoDI.dValorUnitCIFNaMoeda = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitCIFNaMoeda_Col))
        objItemAdicaoDI.dValorUnitFOBEmReal = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitFOBEmReal_Col))
        objItemAdicaoDI.dValorUnitFOBNaMoeda = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitFOBNaMoeda_Col))
        objItemAdicaoDI.sUM = GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
        objItemAdicaoDI.sDescricao = GridItens.TextMatrix(iLinha, iGrid_Descricao_Col)
        objItemAdicaoDI.dIPIUnidadePadraoValor = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_IPIValorUnitario_Col))
             
        objItemAdicaoDI.iAdicao = objAdicaoDI.iSeq
        objItemAdicaoDI.iSeq = iLinha
        
        objAdicaoDI.colItensAdicaoDI.Add objItemAdicaoDI
    
    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr
    
        Case 196654
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_ADICAO_DI_INEXISTENTE", gErr, iItem)
        
        Case 196655

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196656)

    End Select

    Exit Function

End Function

Function Carrega_Moeda()

Dim lErro As Long
Dim objMoeda As ClassMoedas
Dim colMoedas As New Collection

On Error GoTo Erro_Carrega_Moeda
    
    lErro = CF("Moedas_Le_Todas", colMoedas)
    If lErro <> SUCESSO Then gError 196662
    
    'se não existem moedas cadastradas
    If colMoedas.Count = 0 Then gError 196663
    
    For Each objMoeda In colMoedas
    
        If objMoeda.iCodigo <> MOEDA_REAL Then Moeda1.AddItem objMoeda.iCodigo & SEPARADOR & objMoeda.sNome
        If objMoeda.iCodigo <> MOEDA_REAL Then Moeda2.AddItem objMoeda.iCodigo & SEPARADOR & objMoeda.sNome
    
    Next

    Carrega_Moeda = SUCESSO
    
    Exit Function
    
Erro_Carrega_Moeda:

    Carrega_Moeda = gErr
    
    Select Case gErr
    
        Case 196662
        
        Case 196663
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDAS_NAO_CADASTRADAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196664)
    
    End Select

End Function

Private Sub ComparativoMoedaReal_Calcula(ByVal iM As Integer, ByVal dTaxa As Double)

Dim iLinha As Integer, dQtde As Double

    'atualizar o valor da mercadoria, frete e seguro em R$
    If iM = StrParaInt(MoedaMercadoria.Text) Then ValorMercadoriaEmReal.Text = Format(StrParaDbl(ValorMercadoriaMoeda.Text) * dTaxa, "STANDARD")
    If iM = StrParaInt(MoedaFrete.Text) Then ValorFreteInternacEmReal.Text = Format(StrParaDbl(ValorFreteInternacMoeda.Text) * dTaxa, "STANDARD")
    If iM = StrParaInt(MoedaSeguro.Text) Then ValorSeguroInternacEmReal.Text = Format(StrParaDbl(ValorSeguroInternacMoeda.Text) * dTaxa, "STANDARD")
    
    iMercEmRealAlterada = 0
    iFreteEmRealAlterada = 0
    iSeguroEmRealAlterada = 0
    
'    'atualizar o valor FOB em R$ dos itens
'    For iLinha = 1 To objGridItens.iLinhasExistentes
    
'        dQtde = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))
'        GridItens.TextMatrix(iLinha, iGrid_ValorUnitFOBEmReal_Col) = Format(StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitFOBNaMoeda_Col)) * dTaxa, ValorUnitFOBEmReal.Format)
'        GridItens.TextMatrix(iLinha, iGrid_ValorTotalFOBEmReal_Col) = Format(StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorTotalFOBEmReal_Col)) * dTaxa, ValorTotalFOBEmReal.Format)
'        Call Calcula_ValorCIF_Linha(iLinha)
'    Next

    Call RecalcularCIFs
    
End Sub

Private Sub Atualiza_Totais_Itens()

Dim iLinha As Integer
Dim dTotalFOBMoeda As Double, dTotalFOBReal As Double
Dim dTotalCIFMoeda As Double, dTotalCIFReal As Double, dPesoLiq As Double

    'atualizar o valor CIF em R$ dos itens
    For iLinha = 1 To objGridItens.iLinhasExistentes
        
        dTotalFOBMoeda = dTotalFOBMoeda + StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorTotalFOBNaMoeda_Col))
        dTotalFOBReal = dTotalFOBReal + StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorTotalFOBEmReal_Col))
        
        dTotalCIFMoeda = dTotalCIFMoeda + StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorTotalCIFNaMoeda_Col))
        dTotalCIFReal = dTotalCIFReal + StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorTotalCIFEmReal_Col))

        dPesoLiq = dPesoLiq + StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ItemPesoLiq_Col))
        
    Next
    
    LabelTotalFOBDIMoeda.Caption = Format(dTotalFOBMoeda, "standard")
    LabelTotalFOBDIReal.Caption = Format(dTotalFOBReal, "standard")
    LabelTotalCIFDIMoeda.Caption = Format(dTotalCIFMoeda, "standard")
    LabelTotalCIFDIReal.Caption = Format(dTotalCIFReal, "standard")
    LabelPesoLiqDI.Caption = Format(dPesoLiq, "standard")
    
End Sub

Private Sub RecalcularCIFs()

Dim iLinha As Integer

    'atualizar o valor CIF em R$ dos itens
    For iLinha = 1 To objGridItens.iLinhasExistentes
        
        Call Calcula_ValorCIF_Linha(iLinha)
        
    Next
    
    Call Calcula_Valores_Adicao(0, 0)
                
End Sub

Public Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'recolhe o Nome Reduzido da tela
    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Chama a Tela de browse Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

    Exit Sub

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    Fornecedor.Text = objFornecedor.lCodigo
    Call Fornecedor_Validate(bCancel)

    Me.Show

End Sub

Public Sub Fornecedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = REGISTRO_ALTERADO

    Call Fornecedor_Preenche

End Sub

Public Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim bCancel As Boolean

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 1 Then

        If Len(Trim(Fornecedor.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le3(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then gError 196666

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 196667

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

        Case 196666, 196667

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196668)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.ListIndex <> -1 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 196669

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 196670

        sFornecedor = Fornecedor.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 196671

        If lErro = 18272 Then gError 196672

        'coloca na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 196673

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 196669, 196671

        Case 196670
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 196672
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 196673
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196674)

    End Select

    Exit Sub

End Sub

Private Sub Fornecedor_Preenche()
'Reduzido do Fornecedor através da CF Fornecedor_Pesquisa_NomeReduzido em RotinasCPR.ClassCPRSelect'

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 196675

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 196675

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196676)

    End Select
    
    Exit Sub

End Sub

Public Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Filial_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub AdicaoValorAduaneiro_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AdicaoValorAduaneiro_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub AdicaoValorAduaneiro_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub AdicaoValorAduaneiro_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = AdicaoValorAduaneiro
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub IIValor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub IIValor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub IIValor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub IIValor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = IIValor
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub IPIBase_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub IPIBase_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub IPIBase_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub IPIBase_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = IPIBase
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub IPIValor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub IPIValor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub IPIValor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub IPIValor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = IPIValor
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PISBase_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PISBase_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub PISBase_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub PISBase_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = PISBase
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PISAliquota_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PISAliquota_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub PISAliquota_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub PISAliquota_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = PISAliquota
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PISValor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PISValor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub PISValor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub PISValor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = PISValor
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub COFINSBase_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub COFINSBase_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub COFINSBase_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub COFINSBase_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = COFINSBase
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub COFINSAliquota_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub COFINSAliquota_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub COFINSAliquota_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub COFINSAliquota_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = COFINSAliquota
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub COFINSValor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub COFINSValor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub COFINSValor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub COFINSValor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = COFINSValor
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ICMSBase_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ICMSBase_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub ICMSBase_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub ICMSBase_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = ICMSBase
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ICMSValor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ICMSValor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub ICMSValor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub ICMSValor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = ICMSValor
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub AdicaoItem_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AdicaoItem_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub AdicaoItem_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub AdicaoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = AdicaoItem
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ItemPesoBruto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ItemPesoBruto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub ItemPesoBruto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub ItemPesoBruto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ItemPesoBruto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ItemPesoLiq_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ItemPesoLiq_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub ItemPesoLiq_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub ItemPesoLiq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ItemPesoLiq
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub GridDespesas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridDespesas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDespesas, iAlterado)
    End If

End Sub

Private Sub GridDespesas_GotFocus()
    Call Grid_Recebe_Foco(objGridDespesas)
End Sub

Private Sub GridDespesas_EnterCell()
    Call Grid_Entrada_Celula(objGridDespesas, iAlterado)
End Sub

Private Sub GridDespesas_LeaveCell()
    Call Saida_Celula(objGridDespesas)
End Sub

Private Sub GridDespesas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridDespesas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDespesas, iAlterado)
    End If

End Sub

Private Sub GridDespesas_RowColChange()
    Call Grid_RowColChange(objGridDespesas)
End Sub

Private Sub GridDespesas_Scroll()
    Call Grid_Scroll(objGridDespesas)
End Sub

Private Sub GridDespesas_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer

On Error GoTo Erro_GridDespesas_KeyDown

    iLinhasExistentesAnterior = objGridDespesas.iLinhasExistentes
    iItemAtual = GridDespesas.Row

    Call Grid_Trata_Tecla1(KeyCode, objGridDespesas)

    'Se exclui uma linha de itens
    If objGridDespesas.iLinhasExistentes < iLinhasExistentesAnterior Then
    
        'Call Calcula_Valores_Adicao(0, iGrid_ICMSBase_Col)
        Call RecalcularCIFs
    
    End If
    
    Exit Sub
    
Erro_GridDespesas_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196677)

    End Select

    Exit Sub

End Sub

Private Sub GridDespesas_LostFocus()
    Call Grid_Libera_Foco(objGridDespesas)
End Sub

Private Function Inicializa_GridDespesas(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Tipo")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Valor")
    objGrid.colColuna.Add ("%")
    objGrid.colColuna.Add ("Dias")

    'Controles que participam do Grid
    objGrid.colCampo.Add (ComplTipo.Name)
    objGrid.colCampo.Add (ComplDescricao.Name)
    objGrid.colCampo.Add (ComplValor.Name)
    objGrid.colCampo.Add (ComplPerc.Name)
    objGrid.colCampo.Add (ComplDias.Name)

    'Colunas do Grid
    iGrid_ComplTipo_Col = 1
    iGrid_ComplDescricao_Col = 2
    iGrid_ComplValor_Col = 3
    iGrid_ComplPerc_Col = 4
    iGrid_ComplDias_Col = 5

    objGrid.objGrid = GridDespesas

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_LINHAS_GRID_DESPESAS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridDespesas.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridDespesas = SUCESSO

End Function

Private Sub ComplTipo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ComplTipo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDespesas)
End Sub

Private Sub ComplTipo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDespesas)
End Sub

Private Sub ComplTipo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDespesas.objControle = ComplTipo
    lErro = Grid_Campo_Libera_Foco(objGridDespesas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ComplDescricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ComplDescricao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDespesas)
End Sub

Private Sub ComplDescricao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDespesas)
End Sub

Private Sub ComplDescricao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDespesas.objControle = ComplDescricao
    lErro = Grid_Campo_Libera_Foco(objGridDespesas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ComplValor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ComplValor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDespesas)
End Sub

Private Sub ComplValor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDespesas)
End Sub

Private Sub ComplValor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDespesas.objControle = ComplValor
    lErro = Grid_Campo_Libera_Foco(objGridDespesas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ComplDias_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ComplDias_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDespesas)
End Sub

Private Sub ComplDias_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDespesas)
End Sub

Private Sub ComplDias_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDespesas.objControle = ComplDias
    lErro = Grid_Campo_Libera_Foco(objGridDespesas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ComplPerc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ComplPerc_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDespesas)
End Sub

Private Sub ComplPerc_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDespesas)
End Sub

Private Sub ComplPerc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDespesas.objControle = ComplPerc
    lErro = Grid_Campo_Libera_Foco(objGridDespesas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function Saida_Celula_Valor(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula que está deixando de ser a corrente

Dim lErro As Long
Dim bAlterouValor As Boolean
Dim dQtde As Double
Dim dTaxa As Double
Dim dValorAtual As Double

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = objControle
    
    bAlterouValor = False

    If Len(Trim(objControle.Text)) > 0 Then
    
        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 196678

        Select Case objControle.Name
        
            Case ValorUnitFOBNaMoeda.Name
                objControle.Text = Format(objControle.Text, "#,##0.00#####")
            
            Case Else
                objControle.Text = Format(objControle.Text, "STANDARD")
                
        End Select
        
        dValorAtual = StrParaDbl(objControle.Text)
        
        'Se alterou o valor
        If Abs(dValorAtual - StrParaDbl(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, objGridInt.objGrid.Col))) > DELTA_VALORMONETARIO2 Then
            bAlterouValor = True
        End If
        
        If bAlterouValor Then
        
            If objGridInt.objGrid.Name = GridItens.Name Then
            
                Select Case objGridInt.objGrid.Col
                
                    Case iGrid_ValorTotalCIFEmReal_Col
                        GridItens.TextMatrix(GridItens.Row, iGrid_TotalCIFEmRealManual_Col) = "1"
                        Call Grid_Refresh_Checkbox(objGridInt)
                        
                End Select
                
            End If
            
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 196679
        
    If bAlterouValor Then
                        
        If objGridInt.objGrid.Name = GridItens.Name Then
            Call Calcula_ValorCIF_Linha(GridItens.Row)
            Call Calcula_Valores_Adicao(StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_AdicaoItem_Col)), 0)
        Else
            If objGridInt.objGrid.Name = GridAdicao.Name Then
                Call Calcula_Valores_Adicao(GridAdicao.Row, GridAdicao.Col + 1)
            Else
                If objGridInt.objGrid.Name = GridDespesas.Name Then
                
                    'Call Calcula_Valores_Adicao(0, iGrid_ICMSBase_Col)
                    Call RecalcularCIFs

                End If
            End If
        End If
        
    End If

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 196678, 196679
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190678)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Function Saida_Celula_Percentual(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula Data que está deixando de ser a corrente

Dim lErro As Long
Dim dPercent As Double
Dim bAlterouValor As Boolean
Dim dValorAtual As Double

On Error GoTo Erro_Saida_Celula_Percentual

    Set objGridInt.objControle = objControle

    bAlterouValor = False
    dValorAtual = 0
    
    If Len(Trim(objControle.Text)) > 0 Then
    
        'Critica a porcentagem
        lErro = Porcentagem_Critica_Negativa(objControle.Text)
        If lErro <> SUCESSO Then gError 196680

        dPercent = StrParaDbl(objControle.Text)

        'se for igual a 100% -> erro
        If dPercent = 100 Then gError 196681

        objControle.Text = Format(dPercent, "Fixed")
        
        dValorAtual = StrParaDbl(objControle.Text) / 100
        
    End If

    'Se alterou o valor
    If Abs(dValorAtual - PercentParaDbl(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, objGridInt.objGrid.Col))) > DELTA_VALORMONETARIO Then
        bAlterouValor = True
    Else
        bAlterouValor = False
    End If
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 196682

    If bAlterouValor Then
    
        Call Calcula_Valores_Adicao(GridAdicao.Row, GridAdicao.Col)
        
    End If
    
    Saida_Celula_Percentual = SUCESSO

    Exit Function

Erro_Saida_Celula_Percentual:

    Saida_Celula_Percentual = gErr

    Select Case gErr

        Case 196680, 196682
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 196681
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196683)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Function Saida_Celula_Inteiro(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula Data que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Inteiro

    Set objGridInt.objControle = objControle
    
    If Len(Trim(objControle.Text)) > 0 Then
    
        'Critica a porcentagem
        lErro = Valor_Inteiro_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 196680

        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 196682

    
    Saida_Celula_Inteiro = SUCESSO

    Exit Function

Erro_Saida_Celula_Inteiro:

    Saida_Celula_Inteiro = gErr

    Select Case gErr

        Case 196680, 196682
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 196681
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196683)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Padrao(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim bAlterouValor As Boolean

On Error GoTo Erro_Saida_Celula_Padrao

    Set objGridInt.objControle = objControle
    
    bAlterouValor = False
        
    If objGridInt.objGrid.Name = GridItens.Name Then
        
        If objGridInt.objGrid.Col = iGrid_AdicaoItem_Col Then
        
            'Se alterou o item
            If StrParaInt(objControle.Text) <> StrParaInt(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, objGridInt.objGrid.Col)) Then
                bAlterouValor = True
            End If
            
        ElseIf objGridInt.objGrid.Col = iGrid_UnidadeMed_Col Then
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, objGridInt.objGrid.Col) = objControle.Text
        
        End If
        
    End If
       
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 196684
    
    If bAlterouValor Then Call Calcula_Valores_Adicao(0, 0)

    Saida_Celula_Padrao = SUCESSO

    Exit Function

Erro_Saida_Celula_Padrao:

    Saida_Celula_Padrao = gErr

    Select Case gErr

        Case 196684
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196685)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ComplTipo(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objTipoDespesa As New ClassTipoImportCompl

On Error GoTo Erro_Saida_Celula_ComplTipo

    Set objGridInt.objControle = ComplTipo
    
    If Len(Trim(ComplTipo.Text)) > 0 Then
    
        lErro = Valor_Inteiro_Critica(ComplTipo.Text)
        If lErro <> SUCESSO Then gError 196686
        
        objTipoDespesa.iCodigo = StrParaInt(ComplTipo.Text)
        
        lErro = CF("TiposImportCompl_Le", objTipoDespesa)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 196687
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 196688
        
        If objTipoDespesa.iPodeSerDespAduaneira = 0 Then gError 184722
        
         GridDespesas.TextMatrix(GridDespesas.Row, iGrid_ComplDescricao_Col) = objTipoDespesa.sDescReduzida
    
        'verifica se precisa preencher o grid com uma nova linha
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 196689

    Saida_Celula_ComplTipo = SUCESSO

    Exit Function

Erro_Saida_Celula_ComplTipo:

    Saida_Celula_ComplTipo = gErr

    Select Case gErr

        Case 196686, 196687, 196689
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 184722
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOIMPORTCOMPL_NAO_DESPESA", gErr, objTipoDespesa.iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 196688
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOIMPORTCOMPL_NAO_CADASTRADO", gErr, objTipoDespesa.iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196690)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_IPICodigo(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objIPICodigo As New ClassClassificacaoFiscal

On Error GoTo Erro_Saida_Celula_IPICodigo

    Set objGridInt.objControle = IPICodigo

    If Len(Trim(IPICodigo.Text)) > 0 Then
    
        objIPICodigo.sCodigo = IPICodigo.Text
        
        lErro = CF("ClassificacaoFiscal_Le", objIPICodigo)
        If lErro <> SUCESSO And lErro <> 123494 Then gError 196691
    
        If lErro = 123494 Then gError 196692
        
        GridAdicao.TextMatrix(GridAdicao.Row, iGrid_IPIDescricao_Col) = objIPICodigo.sDescricao
        
        GridAdicao.TextMatrix(GridAdicao.Row, iGrid_IIAliquota_Col) = Format(objIPICodigo.dIIAliquota, "Percent")
        
'        GridAdicao.TextMatrix(GridAdicao.Row, iGrid_PISAliquota_Col) = Format(0.0165, "Percent")
'        GridAdicao.TextMatrix(GridAdicao.Row, iGrid_COFINSAliquota_Col) = Format(0.076, "Percent")
    
        GridAdicao.TextMatrix(GridAdicao.Row, iGrid_PISAliquota_Col) = Format(objIPICodigo.dPISAliquota, "Percent")
        GridAdicao.TextMatrix(GridAdicao.Row, iGrid_COFINSAliquota_Col) = Format(objIPICodigo.dCOFINSAliquota, "Percent")
        GridAdicao.TextMatrix(GridAdicao.Row, iGrid_ICMSAliquota_Col) = Format(objIPICodigo.dICMSAliquota, "Percent")
        GridAdicao.TextMatrix(GridAdicao.Row, iGrid_ICMSPercRedBase_Col) = Format(0, "Percent")
        GridAdicao.TextMatrix(GridAdicao.Row, iGrid_IPIAliquota_Col) = Format(objIPICodigo.dIPIAliquota, "Percent")
    
        If (GridAdicao.Row - GridAdicao.FixedRows) = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
        Call Carrega_AdicaoItem
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 196693

    Saida_Celula_IPICodigo = SUCESSO

    Exit Function

Erro_Saida_Celula_IPICodigo:

    Saida_Celula_IPICodigo = gErr

    Select Case gErr

        Case 196691, 196693
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 196692
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CLASSIFICACAOFISCAL_NAO_EXISTENTE", gErr, objIPICodigo.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196694)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim sProduto As String, dQtde As Double, dFator As Double
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto
    
    If Len(Trim(Produto.ClipText)) > 0 Then

        lErro = CF("Produto_Critica2", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 196695
        
        If lErro = 25041 Then gError 196696
        
        'Coloca as demais características do produto na tela
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
            If lErro <> SUCESSO Then gError 196697
            
'            Call StrParaMasked2(Produto, sProduto)
            Produto.PromptInclude = False
            Produto.Text = sProduto
            Produto.PromptInclude = True
        
            'GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
            'If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Descricao_Col))) = 0 Then
            GridItens.TextMatrix(GridItens.Row, iGrid_Descricao_Col) = objProduto.sDescricao
            GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMCompra
        
            dQtde = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
            
            If dQtde > 0 Then
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMCompra, objProduto.sSiglaUMVenda, dFator)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                            
                GridItens.TextMatrix(GridItens.Row, iGrid_ItemPesoBruto_Col) = Formata_Estoque(objProduto.dPesoBruto * dQtde * dFator)
                GridItens.TextMatrix(GridItens.Row, iGrid_ItemPesoLiq_Col) = Formata_Estoque(objProduto.dPesoLiq * dQtde * dFator)
            
                Call Calcula_ValorCIF_Linha(GridItens.Row)
                Call Calcula_Valores_Adicao(StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_AdicaoItem_Col)), 0)
  
            End If
            
            If (GridItens.Row - GridItens.FixedRows) = objGridInt.iLinhasExistentes Then
                objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            End If
            
            
            
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 196698

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 196695, 196698
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 196696
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 196697
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196699)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula Quantidade que está deixando de ser a corrente

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim dQtde As Double, bAlterou As Boolean
Dim iLinha As Integer, dFator As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = objControle

    bAlterou = False
    
    If Len(objControle.Text) > 0 Then
    
        iLinha = GridItens.Row

        lErro = Valor_Positivo_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 196700

        objControle.Text = Formata_Estoque(objControle.Text)
        
        If objGridInt.objGrid.Name = GridItens.Name Then
        
            'Se alterou o valor
            If Abs(StrParaDbl(objControle.Text) - StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))) > QTDE_ESTOQUE_DELTA Then
    
                bAlterou = True
                    
                If objGridInt.objGrid.Col = iGrid_Quantidade_Col Then
            
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
                    If lErro <> SUCESSO Then gError 196701
                   
                    objProduto.sCodigo = sProdutoFormatado
                    
                    'Lê o Produto
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 196702
                    
                    dQtde = StrParaDbl(objControle.Text)
                    
                    GridItens.TextMatrix(iLinha, iGrid_ValorTotalFOBNaMoeda_Col) = Format(StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitFOBNaMoeda_Col)) * dQtde, "STANDARD")
                    GridItens.TextMatrix(iLinha, iGrid_ValorTotalFOBEmReal_Col) = Format(StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitFOBEmReal_Col)) * dQtde, "STANDARD")
                                                                
                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), objProduto.sSiglaUMVenda, dFator)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                                
                    If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_ItemPesoBruto_Col))) = 0 Then GridItens.TextMatrix(iLinha, iGrid_ItemPesoBruto_Col) = Formata_Estoque(objProduto.dPesoBruto * dQtde * dFator)
                    If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_ItemPesoLiq_Col))) = 0 Then GridItens.TextMatrix(iLinha, iGrid_ItemPesoLiq_Col) = Formata_Estoque(objProduto.dPesoLiq * dQtde * dFator)

                End If
                                
            End If
            
        End If
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 196703

    If bAlterou Then
    
        Call Calcula_ValorCIF_Linha(GridItens.Row)
        Call Calcula_Valores_Adicao(StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_AdicaoItem_Col)), 0)
                        
    End If
    
    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 196700 To 196703
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196704)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub BotaoProduto_Click()

Dim lErro As Long
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoProduto_Click

    If Me.ActiveControl Is Produto Then
        sProduto = Produto.Text
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 196705
        
        sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
    End If

    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 196706
    
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
    
    objProduto.sCodigo = sProdutoFormatado
    
    'Lista de produtos produzíveis
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProduto)
    
    Exit Sub

Erro_BotaoProduto_Click:

    Select Case gErr

        Case 196705
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 196706
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196707)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim iLinha As Integer, dQtde As Double, dFator As Double

On Error GoTo Erro_objEventoProduto_evSelecao

    'Verifica se alguma linha está selecionada
    If GridItens.Row < 1 Then gError ERRO_SEM_MENSAGEM

    'Verifica se o Produto está preenchido
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Or Me.ActiveControl Is Produto Then

        Set objProduto = obj1
            
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 196708
            
    '    Call StrParaMasked2(Produto, sProdutoMascarado)
        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True
            
        If Not (Me.ActiveControl Is Produto) Then
            
            GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Descricao_Col))) = 0 Then GridItens.TextMatrix(GridItens.Row, iGrid_Descricao_Col) = objProduto.sDescricao
            GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMCompra
            
            dQtde = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
            
            If dQtde > 0 Then
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMCompra, objProduto.sSiglaUMVenda, dFator)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                            
                GridItens.TextMatrix(GridItens.Row, iGrid_ItemPesoBruto_Col) = Formata_Estoque(objProduto.dPesoBruto * dQtde * dFator)
                GridItens.TextMatrix(GridItens.Row, iGrid_ItemPesoLiq_Col) = Formata_Estoque(objProduto.dPesoLiq * dQtde * dFator)
            
                Call Calcula_ValorCIF_Linha(GridItens.Row)
                Call Calcula_Valores_Adicao(StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_AdicaoItem_Col)), 0)
            
            End If
        
            If (GridItens.Row - GridItens.FixedRows) = objGridItens.iLinhasExistentes Then
                objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
            End If
            
        End If
        
        'Fecha comando de setas se estiver aberto
        Call ComandoSeta_Fechar(Me.Name)
    
    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 196708
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196709)

    End Select

    Exit Sub

End Sub

Public Sub BotaoIPICodigo_Click()

Dim lErro As Long
Dim objIPICodigo As New ClassClassificacaoFiscal
Dim colSelecao As New Collection
Dim sIPICodigo As String

On Error GoTo Erro_BotaoIPICodigo_Click

    If Me.ActiveControl Is IPICodigo Then
        sIPICodigo = IPICodigo.Text
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridAdicao.Row = 0 Then gError 196710
        
        sIPICodigo = GridAdicao.TextMatrix(GridAdicao.Row, iGrid_IPICodigo_Col)
    End If
    
    'Preenche na memória o Código passado
    objIPICodigo.sCodigo = sIPICodigo

    Call Chama_Tela("ClassificacaoFiscalLista", colSelecao, objIPICodigo, objEventoIPICodigo)

    Exit Sub
    
Erro_BotaoIPICodigo_Click:

    Select Case gErr

        Case 196710
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196711)

    End Select

    Exit Sub

End Sub

Private Sub objEventoIPICodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objIPICodigo As New ClassClassificacaoFiscal
Dim bCancel As Boolean
    
On Error GoTo Erro_objEventoIPICodigo_evSelecao
    
    Set objIPICodigo = obj1

    lErro = CF("ClassificacaoFiscal_Le", objIPICodigo)
    If lErro <> SUCESSO And lErro <> 123494 Then gError 196712

    If lErro = 123494 Then gError 196713
    
    IPICodigo.Text = objIPICodigo.sCodigo
    
    If Not (Me.ActiveControl Is IPICodigo) Then

        'Preenche o Cliente com o Cliente selecionado
        GridAdicao.TextMatrix(GridAdicao.Row, iGrid_IPICodigo_Col) = objIPICodigo.sCodigo
        GridAdicao.TextMatrix(GridAdicao.Row, iGrid_IPIDescricao_Col) = objIPICodigo.sDescricao
        
        GridAdicao.TextMatrix(GridAdicao.Row, iGrid_IIAliquota_Col) = Format(objIPICodigo.dIIAliquota, "Percent")
    
        GridAdicao.TextMatrix(GridAdicao.Row, iGrid_PISAliquota_Col) = Format(0.0165, "Percent")
        GridAdicao.TextMatrix(GridAdicao.Row, iGrid_COFINSAliquota_Col) = Format(0.076, "Percent")
        
        If (GridAdicao.Row - GridAdicao.FixedRows) = objGridAdicao.iLinhasExistentes Then
            objGridAdicao.iLinhasExistentes = objGridAdicao.iLinhasExistentes + 1
        End If
        
        Call Carrega_AdicaoItem
    
    End If

    Me.Show

    Exit Sub

Erro_objEventoIPICodigo_evSelecao:

    Select Case gErr
    
        Case 196712
        
        Case 196713
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CLASSIFICACAOFISCAL_NAO_EXISTENTE", gErr, objIPICodigo.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196714)

    End Select

    Exit Sub

End Sub

Private Function Carrega_AdicaoItem() As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Carrega_AdicaoItem

    AdicaoItem.Clear

    For iLinha = 1 To objGridAdicao.iLinhasExistentes
        AdicaoItem.AddItem iLinha
    Next
 
    Exit Function

Erro_Carrega_AdicaoItem:

    Carrega_AdicaoItem = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196715)

    End Select

End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is Produto Then
            Call BotaoProduto_Click
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        ElseIf Me.ActiveControl Is Numero Then
            Call LabelNumero_Click
        ElseIf Me.ActiveControl Is IPICodigo Then
            Call BotaoIPICodigo_Click
        ElseIf Me.ActiveControl Is ComplTipo Then
            Call BotaoTipoDespesa_Click
        ElseIf Me.ActiveControl Is CodigoPC Then
            Call BotaoItensPC_Click
        End If

    End If

End Sub

Public Sub BotaoTipoDespesa_Click()

Dim lErro As Long
Dim objTipoDespesa As New ClassTipoImportCompl
Dim colSelecao As New Collection
Dim sComplTipo As String

On Error GoTo Erro_BotaoTipoDespesa_Click

    If Me.ActiveControl Is ComplTipo Then
        sComplTipo = ComplTipo.Text
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridDespesas.Row = 0 Then gError 196716
        
        sComplTipo = GridDespesas.TextMatrix(GridDespesas.Row, iGrid_ComplTipo_Col)
    End If
    
    'Preenche na memória o Código passado
    objTipoDespesa.iCodigo = StrParaInt(sComplTipo)

    Call Chama_Tela("TiposImportComplLista", colSelecao, objTipoDespesa, objEventoTipoImport, "PodeSerDespAduaneira=1")

    Exit Sub
    
Erro_BotaoTipoDespesa_Click:

    Select Case gErr

        Case 196716
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196717)

    End Select

    Exit Sub

End Sub

Private Sub objEventoTipoImport_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoDespesa As New ClassTipoImportCompl
    
On Error GoTo Erro_objEventoTipoImport_evSelecao
    
    Set objTipoDespesa = obj1
    
    ComplTipo.Text = CStr(objTipoDespesa.iCodigo)

    If Not (Me.ActiveControl Is ComplTipo) Then
    
        'Preenche o Cliente com o Cliente selecionado
        GridDespesas.TextMatrix(GridDespesas.Row, iGrid_ComplTipo_Col) = CStr(objTipoDespesa.iCodigo)
        GridDespesas.TextMatrix(GridDespesas.Row, iGrid_ComplDescricao_Col) = objTipoDespesa.sDescReduzida
    
        'verifica se precisa preencher o grid com uma nova linha
        If objGridDespesas.objGrid.Row - objGridDespesas.objGrid.FixedRows = objGridDespesas.iLinhasExistentes Then
            objGridDespesas.iLinhasExistentes = objGridDespesas.iLinhasExistentes + 1
        End If
        
    End If
    
    Me.Show

    Exit Sub

Erro_objEventoTipoImport_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196718)

    End Select

    Exit Sub

End Sub

Private Function Calcula_Valores() As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer
Dim dDespesas As Double
Dim dII As Double
Dim dIPI As Double
Dim dPIS As Double
Dim dCOFINS As Double
Dim dICMS As Double

On Error GoTo Erro_Calcula_Valores

    Call Calcula_Valores_Despesas
    
    For iLinha = 1 To objGridAdicao.iLinhasExistentes
        dII = dII + StrParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_IIValor_Col))
        dIPI = dIPI + StrParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_IPIValor_Col))
        dPIS = dPIS + StrParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_PISValor_Col))
        dCOFINS = dCOFINS + StrParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_COFINSValor_Col))
        dICMS = dICMS + StrParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_ICMSValor_Col))
    Next
    
    DIIIValor.Text = Format(dII, "STANDARD")
    DIIPIValor.Text = Format(dIPI, "STANDARD")
    DIPISValor.Text = Format(dPIS, "STANDARD")
    DICOFINSValor.Text = Format(dCOFINS, "STANDARD")
    DIICMSValor.Text = Format(dICMS, "STANDARD")
    
    DIValorFOB.Text = ValorMercadoriaEmReal.Text
    DIValorFrete.Text = ValorFreteInternacEmReal.Text
    DIValorSeguro.Text = ValorSeguroInternacEmReal.Text
    DIValorCIF.Text = Arredonda_Moeda(StrParaDbl(LabelTotalCIFDIReal.Caption))
    
    DIValorProdutos.Text = Arredonda_Moeda(StrParaDbl(DIValorCIF.Text) + dII)
    DIValorOutrasDesp.Text = Arredonda_Moeda(dPIS + dCOFINS + StrParaDbl(DIValorDespesas.Text))
    DIValorTotal.Text = Arredonda_Moeda(StrParaDbl(DIValorProdutos.Text) + StrParaDbl(DIValorOutrasDesp.Text) + dIPI + dICMS)
    
    Calcula_Valores = SUCESSO

    Exit Function

Erro_Calcula_Valores:

    Calcula_Valores = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196719)

    End Select

End Function

Private Function Calcula_Valores_Despesas() As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer, objTipoDespesa As New ClassTipoImportCompl
Dim dDespesas As Double, dDespesasICMS As Double

On Error GoTo Erro_Calcula_Valores_Despesas

    For iLinha = 1 To objGridDespesas.iLinhasExistentes
    
        objTipoDespesa.iCodigo = StrParaInt(GridDespesas.TextMatrix(iLinha, iGrid_ComplTipo_Col))
        
        lErro = CF("TiposImportCompl_Le", objTipoDespesa)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 196687
        If lErro = SUCESSO Then
        
            If objTipoDespesa.iIncluiBaseICMS = 0 Then
                dDespesas = dDespesas + StrParaDbl(GridDespesas.TextMatrix(iLinha, iGrid_ComplValor_Col))
            Else
                dDespesasICMS = dDespesasICMS + StrParaDbl(GridDespesas.TextMatrix(iLinha, iGrid_ComplValor_Col))
            End If
            
        End If
    Next
    
    LabelDespICMS.Caption = Format(dDespesasICMS, "STANDARD")
    LabelOutrasDesp.Caption = Format(dDespesas, "STANDARD")
    DIValorDespesas.Text = Format(dDespesasICMS, "STANDARD")
    
    Calcula_Valores_Despesas = SUCESSO

    Exit Function

Erro_Calcula_Valores_Despesas:

    Calcula_Valores_Despesas = gErr

    Select Case gErr

        Case 196687
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196720)

    End Select

End Function

Private Function Calcula_Valores_Adicao(ByVal iLinhaAdicao As Integer, ByVal iColunaInicial As Integer) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer, dICMSValor As Double
Dim iLinhaAux As Integer, dICMSBase As Double, dValorAduaneiroDI As Double
Dim dValorAduaneiro As Double, dIIAliquota As Double, dIIValor As Double
Dim dIPIBase As Double, dIPIAliquota As Double, dIPIValor As Double, dICMSAliquota As Double, dICMSPercRedBase As Double
Dim dPISValor As Double, dCOFINSValor As Double, dPISAliquota As Double, dCOFINSAliquota As Double
Dim dPISCOFINSBase As Double, objTipoDespesa As New ClassTipoImportCompl
Dim dValorDespesasICMS As Double, objDespesas As New ClassImportCompl
Dim dDespAduaneira As Double, dDespAduaneiraDI As Double, dTaxaSiscomexDI As Double, dTaxaSiscomex As Double
Dim dICMSAliquotaEfetiva As Double, dIPIValorUnitAcum As Double, W As Double, Y As Double
Dim dtDataDI As Date 'para saber qual a regra do calculo do pis/cofins

On Error GoTo Erro_Calcula_Valores_Adicao

    dtDataDI = StrParaDate(Data.Text)

    'obter total das despesas aduaneiras que irá para a base de calculo do icms
    For iLinha = 1 To objGridDespesas.iLinhasExistentes
    
        objDespesas.iTipo = StrParaInt(GridDespesas.TextMatrix(iLinha, iGrid_ComplTipo_Col))
    
        If objDespesas.iTipo <> 0 Then
        
            objTipoDespesa.iCodigo = objDespesas.iTipo
            
            lErro = CF("TiposImportCompl_Le", objTipoDespesa)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 196687
            
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 196688
        
            objDespesas.dValor = StrParaDbl(GridDespesas.TextMatrix(iLinha, iGrid_ComplValor_Col))
            
            If objTipoDespesa.iIncluiBaseICMS <> 0 Then
                dValorDespesasICMS = dValorDespesasICMS + objDespesas.dValor
                If objTipoDespesa.iCodigo = 4 Then
                    dTaxaSiscomexDI = dTaxaSiscomexDI + objDespesas.dValor
                Else
                    dDespAduaneiraDI = dDespAduaneiraDI + objDespesas.dValor
                End If
            End If
    
        End If
        
    Next
    
    dValorAduaneiroDI = 0
    For iLinha = 1 To objGridItens.iLinhasExistentes
        dValorAduaneiroDI = Arredonda_Moeda(dValorAduaneiroDI + StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorTotalCIFEmReal_Col)))
    Next
    
    For iLinhaAux = 1 To objGridAdicao.iLinhasExistentes
        
        If iLinhaAdicao = 0 Or iLinhaAdicao = iLinhaAux Then
        
            dValorAduaneiro = 0
            dIPIValorUnitAcum = 0
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If iLinhaAux = StrParaInt(GridItens.TextMatrix(iLinha, iGrid_AdicaoItem_Col)) Then
                    dValorAduaneiro = dValorAduaneiro + StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorTotalCIFEmReal_Col))
                    dIPIValorUnitAcum = dIPIValorUnitAcum + (StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_IPIValorUnitario_Col)) * StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)))
                End If
            Next
            GridAdicao.TextMatrix(iLinhaAux, iGrid_AdicaoValorAduaneiro_Col) = Format(dValorAduaneiro, "STANDARD")
            
            If iColunaInicial <= iGrid_DespAdua_Col Then
                If dValorAduaneiroDI <> 0 Then GridAdicao.TextMatrix(iLinhaAux, iGrid_DespAdua_Col) = Format(dDespAduaneiraDI * (dValorAduaneiro / dValorAduaneiroDI), "STANDARD")
            End If
            dDespAduaneira = StrParaDbl(GridAdicao.TextMatrix(iLinhaAux, iGrid_DespAdua_Col))
'            If dValorAduaneiroDI <> 0 Then
'                dDespAduaneira = dValorAduaneiro / dValorAduaneiroDI * dDespAduaneiraDI
'            Else
'                dDespAduaneira = 0
'            End If
            
            If iColunaInicial <= iGrid_TaxaSiscomex_Col Then
                Call Calcula_Siscomex_Adicao(iLinhaAux, dTaxaSiscomexDI, dTaxaSiscomex)
                GridAdicao.TextMatrix(iLinhaAux, iGrid_TaxaSiscomex_Col) = Format(dTaxaSiscomex, "STANDARD")
            End If
            dTaxaSiscomex = StrParaDbl(GridAdicao.TextMatrix(iLinhaAux, iGrid_TaxaSiscomex_Col))
            
            'dValorDespesasICMS = dDespAduaneira + dTaxaSiscomex
            
            dIIAliquota = PercentParaDbl(GridAdicao.TextMatrix(iLinhaAux, iGrid_IIAliquota_Col))
            If iColunaInicial <= iGrid_IIValor_Col Then
                dIIValor = Arredonda_Moeda(dIIAliquota * dValorAduaneiro)
                GridAdicao.TextMatrix(iLinhaAux, iGrid_IIValor_Col) = Format(dIIValor, "STANDARD")
            Else
                dIIValor = StrParaDbl(GridAdicao.TextMatrix(iLinhaAux, iGrid_IIValor_Col))
            End If
            
            If iColunaInicial <= iGrid_IPIBase_Col Then
                dIPIBase = Arredonda_Moeda(dValorAduaneiro + dIIValor)
                GridAdicao.TextMatrix(iLinhaAux, iGrid_IPIBase_Col) = Format(dIPIBase, "STANDARD")
            Else
                dIPIBase = StrParaDbl(GridAdicao.TextMatrix(iLinhaAux, iGrid_IPIBase_Col))
            End If
            
            dIPIAliquota = PercentParaDbl(GridAdicao.TextMatrix(iLinhaAux, iGrid_IPIAliquota_Col))
            
            If iColunaInicial <= iGrid_IPIValor_Col Then
                If dIPIValorUnitAcum = 0 Then
                    dIPIValor = Arredonda_Moeda(dIPIBase * dIPIAliquota)
                Else
                    dIPIValor = dIPIValorUnitAcum
                End If
                GridAdicao.TextMatrix(iLinhaAux, iGrid_IPIValor_Col) = Format(dIPIValor, "STANDARD")
            Else
                dIPIValor = StrParaDbl(GridAdicao.TextMatrix(iLinhaAux, iGrid_IPIValor_Col))
            End If
            
            dICMSAliquota = PercentParaDbl(GridAdicao.TextMatrix(iLinhaAux, iGrid_ICMSAliquota_Col))
            dICMSPercRedBase = PercentParaDbl(GridAdicao.TextMatrix(iLinhaAux, iGrid_ICMSPercRedBase_Col))
            dICMSAliquotaEfetiva = Round(dICMSAliquota * (1 - dICMSPercRedBase), 4)
            dPISAliquota = PercentParaDbl(GridAdicao.TextMatrix(iLinhaAux, iGrid_PISAliquota_Col))
            dCOFINSAliquota = PercentParaDbl(GridAdicao.TextMatrix(iLinhaAux, iGrid_COFINSAliquota_Col))
            
            If dtDataDI >= DATA_PIS_NOVO_CALC Or dtDataDI = DATA_NULA Then
            
                dPISCOFINSBase = dValorAduaneiro
            
            Else
            
                'vide http://www.receita.fazenda.gov.br/Legislacao/Ins/2005/in5722005.htm
                If dIPIValorUnitAcum = 0 Then
                    dPISCOFINSBase = dValorAduaneiro * (1 + (dICMSAliquotaEfetiva * (dIIAliquota + (dIPIAliquota * (1 + dIIAliquota))))) / ((1 - dPISAliquota - dCOFINSAliquota) * (1 - dICMSAliquotaEfetiva))
                Else
                    Y = (1 + (dICMSAliquotaEfetiva * dIIAliquota)) / ((1 - dPISAliquota - dCOFINSAliquota) * (1 - dICMSAliquotaEfetiva))
                    dPISCOFINSBase = (dValorAduaneiro * Y)
                    For iLinha = 1 To objGridItens.iLinhasExistentes
                        If iLinhaAux = StrParaInt(GridItens.TextMatrix(iLinha, iGrid_AdicaoItem_Col)) Then
                            W = (dICMSAliquotaEfetiva * StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_IPIValorUnitario_Col))) / ((1 - dPISAliquota - dCOFINSAliquota) * (1 - dICMSAliquotaEfetiva))
                            dPISCOFINSBase = dPISCOFINSBase + (W * StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)))
                        End If
                    Next
                    
                End If
            
            End If
            
            GridAdicao.TextMatrix(iLinhaAux, iGrid_PISBase_Col) = Format(dPISCOFINSBase, "STANDARD")
            GridAdicao.TextMatrix(iLinhaAux, iGrid_COFINSBase_Col) = Format(dPISCOFINSBase, "STANDARD")
                    
            If iColunaInicial <= iGrid_PISValor_Col Then
                dPISValor = Arredonda_Moeda(dPISAliquota * dPISCOFINSBase)
                GridAdicao.TextMatrix(iLinhaAux, iGrid_PISValor_Col) = Format(dPISValor, "STANDARD")
            Else
                dPISValor = StrParaDbl(GridAdicao.TextMatrix(iLinhaAux, iGrid_PISValor_Col))
            End If
            
            If iColunaInicial <= iGrid_COFINSValor_Col Then
                dCOFINSValor = Arredonda_Moeda(dCOFINSAliquota * dPISCOFINSBase)
                GridAdicao.TextMatrix(iLinhaAux, iGrid_COFINSValor_Col) = Format(dCOFINSValor, "STANDARD")
            Else
                dCOFINSValor = StrParaDbl(GridAdicao.TextMatrix(iLinhaAux, iGrid_COFINSValor_Col))
            End If
            
            If iColunaInicial <= iGrid_ICMSBase_Col Then
                dICMSBase = dValorAduaneiro + dIIValor + dIPIValor + dPISValor + dCOFINSValor + dTaxaSiscomex + dDespAduaneira
                If gobjCRFAT.iNFImportacaoTribFlag09 = DESMARCADO Then
                    dICMSBase = dICMSBase / (1 - dICMSAliquotaEfetiva)
                Else
                    dICMSBase = (dICMSBase * (1 - dICMSPercRedBase)) / (1 - dICMSAliquota)
                End If
                GridAdicao.TextMatrix(iLinhaAux, iGrid_ICMSBase_Col) = Format(dICMSBase, "STANDARD")
            Else
                dICMSBase = StrParaDbl(GridAdicao.TextMatrix(iLinhaAux, iGrid_ICMSBase_Col))
            End If
            
            If iColunaInicial <= iGrid_ICMSValor_Col Then
                If gobjCRFAT.iNFImportacaoTribFlag09 = DESMARCADO Then
                    dICMSValor = Arredonda_Moeda(dICMSBase * dICMSAliquotaEfetiva)
                Else
                    dICMSValor = Arredonda_Moeda(dICMSBase * dICMSAliquota)
                End If
                GridAdicao.TextMatrix(iLinhaAux, iGrid_ICMSValor_Col) = Format(dICMSValor, "STANDARD")
            Else
                dICMSValor = StrParaDbl(GridAdicao.TextMatrix(iLinhaAux, iGrid_ICMSValor_Col))
            End If
        
        End If
        
    Next
    
    Call Calcula_Valores
                    
    Calcula_Valores_Adicao = SUCESSO

    Exit Function

Erro_Calcula_Valores_Adicao:

    Calcula_Valores_Adicao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196721)

    End Select

End Function

Private Function Calcula_ValorCIF_Linha(ByVal iLinha As Integer) As Long
'recalcula valor cif

Dim lErro As Long, dValorTotalCIFEmRealNoGrid As Double, iTotalCIFEmRealManual As Integer, iLinha2 As Integer
Dim dValorUnitCIFEmReal As Double, dValorUnitFOBEmReal As Double, dValorTotalFOBEmReal As Double, dValorTotalCIFEmReal As Double
Dim dValorUnitCIFNaMoeda As Double, dValorUnitFobMoeda As Double, dValorTotalFobMoeda As Double, dValorTotalCIFMoeda As Double
Dim dFatorPeso As Double, dPesoLiq As Double, dQtde As Double, dTaxaItens As Double, dDespAduaneirasPeso As Double, objTipoDespesa As New ClassTipoImportCompl
Dim dTaxaFrete As Double, dTaxaSeg As Double, dFatorValor As Double, dDespAduaneirasValor As Double

On Error GoTo Erro_Calcula_ValorCIF_Linha
    
    dValorTotalCIFEmRealNoGrid = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorTotalCIFEmReal_Col))
    dValorTotalFOBEmReal = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorTotalFOBEmReal_Col))
    dValorUnitFobMoeda = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitFOBNaMoeda_Col))
    dPesoLiq = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ItemPesoLiq_Col))
    dQtde = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))
    iTotalCIFEmRealManual = StrParaInt(GridItens.TextMatrix(iLinha, iGrid_TotalCIFEmRealManual_Col))
    
    If StrParaInt(MoedaItens.Text) = 1 Then
        dTaxaItens = StrParaDbl(TaxaMoeda1.Text)
    Else
        dTaxaItens = StrParaDbl(TaxaMoeda2.Text)
    End If

    If StrParaInt(MoedaFrete.Text) = 1 Then
        dTaxaFrete = StrParaDbl(TaxaMoeda1.Text)
    Else
        dTaxaFrete = StrParaDbl(TaxaMoeda2.Text)
    End If
    
    If StrParaInt(MoedaSeguro.Text) = 1 Then
        dTaxaSeg = StrParaDbl(TaxaMoeda1.Text)
    Else
        dTaxaSeg = StrParaDbl(TaxaMoeda2.Text)
    End If
    
    'para evitar divisao por zero
    If dQtde = 0 Then dQtde = 1
    
    If StrParaDbl(DIPesoLiquido.Text) > 0 Then
        dFatorPeso = dPesoLiq / StrParaDbl(DIPesoLiquido.Text)
    Else
        dFatorPeso = 1
    End If
    
    dValorTotalFobMoeda = Arredonda_Moeda(dValorUnitFobMoeda * dQtde, 7)
    GridItens.TextMatrix(iLinha, iGrid_ValorTotalFOBNaMoeda_Col) = Format(dValorTotalFobMoeda, "#,##0.00#####")
    GridItens.TextMatrix(iLinha, iGrid_ValorUnitFOBEmReal_Col) = Format(dValorUnitFobMoeda * dTaxaItens, "STANDARD")
    
    dDespAduaneirasPeso = 0
    dDespAduaneirasValor = 0
    
    For iLinha2 = 1 To objGridDespesas.iLinhasExistentes
    
        objTipoDespesa.iCodigo = StrParaInt(GridDespesas.TextMatrix(iLinha2, iGrid_ComplTipo_Col))
        
        If objTipoDespesa.iCodigo <> 0 Then
        
            lErro = CF("TiposImportCompl_Le", objTipoDespesa)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 196687
            If lErro = SUCESSO Then
            
                If objTipoDespesa.iIncluiNoValorAduaneiro = 1 Then
                
                    Select Case objTipoDespesa.iTipoRateio
                    
                        Case IMPORTCOMPL_TIPORATEIO_PESO
                            dDespAduaneirasPeso = dDespAduaneirasPeso + StrParaDbl(GridDespesas.TextMatrix(iLinha2, iGrid_ComplValor_Col))
                            
                        Case IMPORTCOMPL_TIPORATEIO_VALOR
                            dDespAduaneirasValor = dDespAduaneirasValor + StrParaDbl(GridDespesas.TextMatrix(iLinha2, iGrid_ComplValor_Col))
                    
                    End Select
                    
                End If
                
            End If
        End If
    Next
    
    If Abs(StrParaDbl(ValorMercadoriaEmReal.Text)) > DELTA_VALORMONETARIO Then
        dFatorValor = dValorTotalFOBEmReal / StrParaDbl(ValorMercadoriaEmReal.Text)
    Else
        dFatorValor = dFatorPeso
    End If
    
    dValorTotalCIFMoeda = dValorTotalFobMoeda + ((StrParaDbl(ValorSeguroInternacMoeda.Text) * dTaxaSeg / dTaxaItens) * dFatorValor) + ((StrParaDbl(ValorFreteInternacMoeda.Text) * dTaxaFrete / dTaxaItens) * dFatorPeso)
    dValorUnitCIFNaMoeda = dValorTotalCIFMoeda / dQtde
    
    If iTotalCIFEmRealManual = 0 Then
        GridItens.TextMatrix(iLinha, iGrid_ValorTotalFOBEmReal_Col) = Format(dValorTotalFobMoeda * dTaxaItens, "STANDARD")
        dValorTotalCIFEmReal = Arredonda_Moeda((dValorTotalCIFMoeda * dTaxaItens) + (dDespAduaneirasPeso * dFatorPeso) + (dDespAduaneirasValor * dFatorValor))
        GridItens.TextMatrix(iLinha, iGrid_ValorTotalCIFEmReal_Col) = Format(dValorTotalCIFEmReal, ValorTotalCIFEmReal.Format)
    Else
        dValorTotalCIFEmReal = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorTotalCIFEmReal_Col))
    End If
    
    dValorUnitCIFEmReal = dValorTotalCIFEmReal / dQtde
    
    GridItens.TextMatrix(iLinha, iGrid_ValorUnitCIFEmReal_Col) = Format(dValorUnitCIFEmReal, "STANDARD")
    GridItens.TextMatrix(iLinha, iGrid_ValorUnitCIFNaMoeda_Col) = Format(dValorUnitCIFNaMoeda, "STANDARD")
    GridItens.TextMatrix(iLinha, iGrid_ValorTotalCIFNaMoeda_Col) = Format(dValorTotalCIFMoeda, "#,##0.00#####")

    Call Atualiza_Totais_Adicao(iLinha)
    Call Atualiza_Totais_Itens
        
    Calcula_ValorCIF_Linha = SUCESSO

    Exit Function

Erro_Calcula_ValorCIF_Linha:

    Calcula_ValorCIF_Linha = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196722)

    End Select

End Function

Private Sub DIPesoBruto_Change()
    iAlterado = REGISTRO_ALTERADO
    iDIPesoBrutoAlterada = REGISTRO_ALTERADO
End Sub

Private Sub DIPesoBruto_Validate(Cancel As Boolean)

Dim lErro As Long
    
On Error GoTo Erro_DIPesoBruto_Validate

    If iDIPesoBrutoAlterada <> 0 Then
    
        'Criticao valor
        lErro = Valor_NaoNegativo_Critica(DIPesoBruto.Text)
        If lErro <> SUCESSO Then gError 196723
    
        'Coloca o valor formatado na Tela
        DIPesoBruto.Text = Formata_Estoque(DIPesoBruto.Text)
    
        Call RecalcularCIFs
        
        iDIPesoBrutoAlterada = 0
    
    End If
    
    Exit Sub

Erro_DIPesoBruto_Validate:

    Cancel = True

    Select Case gErr
    
        Case 196723

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196724)

    End Select

End Sub

Private Sub DIPesoLiquido_Change()
    iAlterado = REGISTRO_ALTERADO
    iDIPesoLiqAlterada = REGISTRO_ALTERADO
End Sub

Private Sub DIPesoliquido_Validate(Cancel As Boolean)

Dim lErro As Long
    
On Error GoTo Erro_DIPesoliquido_Validate

    If iDIPesoLiqAlterada <> 0 Then
    
        'Criticao valor
        lErro = Valor_Positivo_Critica(DIPesoLiquido.Text)
        If lErro <> SUCESSO Then gError 196725
    
        'Coloca o valor formatado na Tela
        DIPesoLiquido.Text = Formata_Estoque(DIPesoLiquido.Text)
    
        Call RecalcularCIFs
    
        iDIPesoLiqAlterada = 0
    
    End If
    
    Exit Sub

Erro_DIPesoliquido_Validate:

    Cancel = True

    Select Case gErr
    
        Case 196725

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196726)

    End Select

End Sub

Private Sub ZerarFlagsAlteracao()

    iDataAlterada = 0

    iTaxaMoeda1Alterada = 0
    iTaxaMoeda2Alterada = 0
    
    iMercEmMoedaAlterada = 0
    iFreteEmMoedaAlterada = 0
    iSeguroEmMoedaAlterada = 0
    
    iMercEmRealAlterada = 0
    iFreteEmRealAlterada = 0
    iSeguroEmRealAlterada = 0
    
    iDIPesoBrutoAlterada = 0
    iDIPesoLiqAlterada = 0
       
    iMoedaMercadoriaAnt = StrParaInt(MoedaMercadoria.Text)
    iMoedaFreteAnt = StrParaInt(MoedaFrete.Text)
    iMoedaSeguroAnt = StrParaInt(MoedaSeguro.Text)
    iMoedaItensAnt = StrParaInt(MoedaItens.Text)
   
End Sub

Private Function DIInfo_Critica(ByVal objDIInfo As ClassDIInfo) As Long

Dim lErro As Long, dValorFOBTotalEmReal As Double, dValorCIFTotalEmReal As Double
Dim objAdicaoDI As ClassAdicaoDI, dDespAduaneiras As Double, iLinha2 As Integer
Dim objItemAdicaoDI As ClassItemAdicaoDI, objTipoDespesa As New ClassTipoImportCompl
Dim vbResult As VbMsgBoxResult, dTaxaSiscomexDesp As Double, dTaxaSiscomexAdicao As Double
Dim objItemPCDI As ClassItemPCDI
Dim objPedidoCompra As New ClassPedidoCompras
Dim objItemPC As ClassItemPedCompra
Dim dQuantDisponivel As Double
Dim dQuantPC As Double
Dim objProd As ClassProduto
Dim dFator As Double
Dim dQuantAdicaoDIProd As Double
Dim objItemPCDI_1 As ClassItemPCDI
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim iAchou As Integer

On Error GoTo Erro_DIInfo_Critica
 
    dDespAduaneiras = 0
    For iLinha2 = 1 To objGridDespesas.iLinhasExistentes
    
        objTipoDespesa.iCodigo = StrParaInt(GridDespesas.TextMatrix(iLinha2, iGrid_ComplTipo_Col))
        
        If objTipoDespesa.iCodigo <> 0 Then
        
            lErro = CF("TiposImportCompl_Le", objTipoDespesa)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 196687
            If lErro = SUCESSO Then
            
                If objTipoDespesa.iIncluiNoValorAduaneiro = 1 Then
                    dDespAduaneiras = dDespAduaneiras + StrParaDbl(GridDespesas.TextMatrix(iLinha2, iGrid_ComplValor_Col))
                End If
                If objTipoDespesa.iIncluiBaseICMS = MARCADO Then
                    If objTipoDespesa.iCodigo = 4 Then
                        dTaxaSiscomexDesp = dTaxaSiscomexDesp + StrParaDbl(GridDespesas.TextMatrix(iLinha2, iGrid_ComplValor_Col))
                    End If
                End If
            End If
        End If
    Next
    
    For Each objAdicaoDI In objDIInfo.colAdicoesDI
    
        For Each objItemAdicaoDI In objAdicaoDI.colItensAdicaoDI
        
            dValorFOBTotalEmReal = dValorFOBTotalEmReal + objItemAdicaoDI.dValorTotalFOBEmReal
            dValorCIFTotalEmReal = dValorCIFTotalEmReal + objItemAdicaoDI.dValorTotalCIFEmReal
        
        Next
        
        dTaxaSiscomexAdicao = dTaxaSiscomexAdicao + objAdicaoDI.dTaxaSiscomex
        
    Next
    
    'o valor fob tem que ser o valor da mercadoria
    If Abs(dValorFOBTotalEmReal - objDIInfo.dValorMercadoriaEmReal) > DELTA_VALORMONETARIO Then 'gError 184741
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_DIINFO_TOTAL_FOB_EM_REAL", Format(dValorFOBTotalEmReal, ValorMercadoriaEmReal.Format), Format(objDIInfo.dValorMercadoriaEmReal, ValorMercadoriaEmReal.Format))
        If vbResult = vbNo Then gError 184741
    End If
    
    'o valor cif tem que ser o fob mais o frete e seguro internacional
    If Abs(dValorCIFTotalEmReal - (dDespAduaneiras + objDIInfo.dValorMercadoriaEmReal + objDIInfo.dValorFreteInternacEmReal + objDIInfo.dValorSeguroInternacEmReal)) > DELTA_VALORMONETARIO Then 'gError 184742
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_DIINFO_TOTAL_CIF_EM_REAL", Format(dValorCIFTotalEmReal, ValorMercadoriaEmReal.Format), Format(dDespAduaneiras + objDIInfo.dValorMercadoriaEmReal + objDIInfo.dValorFreteInternacEmReal + objDIInfo.dValorSeguroInternacEmReal, ValorMercadoriaEmReal.Format))
        If vbResult = vbNo Then gError 184742
    End If
    
    'o valor das siscomex da adição não bate com o da despesa
    If Abs(dTaxaSiscomexDesp - dTaxaSiscomexAdicao) > DELTA_VALORMONETARIO Then
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_DIINFO_TOTAL_TAXASISCOMEX", Format(dTaxaSiscomexDesp, ValorMercadoriaEmReal.Format), Format(dTaxaSiscomexAdicao, ValorMercadoriaEmReal.Format))
        If vbResult = vbNo Then gError 184742
    End If

    For Each objItemPCDI In objDIInfo.colItensPC
        
        Set objPedidoCompra = New ClassPedidoCompras
                
        objPedidoCompra.lCodigo = objItemPCDI.lCodigoPC
        objPedidoCompra.iFilialEmpresa = giFilialEmpresa

        'Lê os itens de um Pedido de Compra a partir do código e de FilialEmpresa do Pedido de Compras
        lErro = CF("ItensPC_Le_Codigo", objPedidoCompra)
        If lErro <> SUCESSO And lErro <> 25605 Then gError 210588

        dQuantDisponivel = 0

        For Each objItemPC In objPedidoCompra.colItens

            If objItemPC.sProduto = objItemPCDI.sProdutoPC Then
                dQuantDisponivel = objItemPC.dQuantidade - objItemPC.dQuantRecebida - objItemPC.dQuantRecebimento
                Exit For
            End If
                
        Next
        
        'le a quantidade do produto de ItensPC que esta em DI e que  esta associada a nota fiscal
        lErro = CF("ItensPCDI_Le_QuantPC", objDIInfo.sNumero, objItemPCDI, dQuantPC)
        If lErro <> SUCESSO Then gError 210619
        
        dQuantDisponivel = dQuantDisponivel - dQuantPC
        
        If objItemPCDI.dQuantPC > dQuantDisponivel Then
            vbResult = Rotina_Aviso(vbYesNo, "AVISO_QUANTPC_MAIOR_DISPONIVEL", objItemPCDI.lCodigoPC, objItemPCDI.sProdutoPC, objItemPCDI.dQuantPC, dQuantDisponivel)
            If vbResult = vbNo Then gError 210620
        End If
        
    Next
    
    iIndice = 0
    
    'soma os produtos de colItensPC com o mesmo codigo de produto
    For Each objItemPCDI In objDIInfo.colItensPC
    
        iIndice = iIndice + 1
                
        Set objProd = New ClassProduto

        objProd.sCodigo = objItemPCDI.sProdutoPC

        lErro = CF("Produto_Le", objProd)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
                
        dQuantPC = 0
                
        iAchou = 1
                
        For iIndice1 = 1 To objDIInfo.colItensPC.Count
    
            Set objItemPCDI_1 = objDIInfo.colItensPC(iIndice1)
    
            If objItemPCDI.sProdutoPC = objItemPCDI_1.sProdutoPC Then
                If iIndice1 < iIndice Then
                    iAchou = 0
                    Exit For
                End If
        
                lErro = CF("UM_Conversao_Trans", objProd.iClasseUM, objItemPCDI_1.sUMPC, objProd.sSiglaUMEstoque, dFator)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                dQuantPC = dQuantPC + (objItemPCDI_1.dQuantPC * dFator)
                
            End If
    
        Next
    
        
        'soma os ItensAdicaoDI com o mesmo codigo de produto
        If iAchou = 1 Then
        
            dQuantAdicaoDIProd = 0
            
            For Each objAdicaoDI In objDIInfo.colAdicoesDI
            
                For Each objItemAdicaoDI In objAdicaoDI.colItensAdicaoDI
                
                    If objItemAdicaoDI.sProduto = objItemPCDI.sProdutoPC Then
                
                        lErro = CF("UM_Conversao_Trans", objProd.iClasseUM, objItemAdicaoDI.sUM, objProd.sSiglaUMEstoque, dFator)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
                        dQuantAdicaoDIProd = dQuantAdicaoDIProd + (objItemAdicaoDI.dQuantidade * dFator)
                        
                    End If
                
                Next
                
            Next
            
            'verifica se a quantidade do produto em ItensPC difere da quantidade em AdicaoDI
            If dQuantAdicaoDIProd <> dQuantPC Then
                vbResult = Rotina_Aviso(vbYesNo, "AVISO_QUANT_PROD_ADICAODI_DIFERE_ITEMPC", objItemPCDI.lCodigoPC, objItemPCDI.sProdutoPC, dQuantPC, dQuantAdicaoDIProd)
                If vbResult = vbNo Then gError 210620
            End If
    
        End If
    
    Next
    
    DIInfo_Critica = SUCESSO
    
    Exit Function
    
Erro_DIInfo_Critica:

    DIInfo_Critica = gErr

    Select Case gErr

        Case 184741
            'Call Rotina_Erro(vbOKOnly, "ERRO_DIINFO_TOTAL_FOB_EM_REAL", gErr)
        
        Case 184742
            'Call Rotina_Erro(vbOKOnly, "ERRO_DIINFO_TOTAL_CIF_EM_REAL", gErr)
        
        Case 210620, ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184740)

    End Select
    
    Exit Function

End Function

Private Sub Atualiza_Totais_Adicao(ByVal iLinhaAdicao As Integer)

Dim iLinha As Integer, iItem As Integer, iAdicao As Integer
Dim dTotalFOBMoeda As Double, dTotalFOBReal As Double
Dim dTotalCIFMoeda As Double, dTotalCIFReal As Double, dPesoLiq As Double

    iAdicao = StrParaInt(GridItens.TextMatrix(iLinhaAdicao, iGrid_AdicaoItem_Col))
    For iLinha = 1 To objGridItens.iLinhasExistentes
    
        iItem = StrParaInt(GridItens.TextMatrix(iLinha, iGrid_AdicaoItem_Col))
        If iItem = iAdicao Then
        
            dTotalFOBMoeda = dTotalFOBMoeda + StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorTotalFOBNaMoeda_Col))
            dTotalFOBReal = dTotalFOBReal + StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorTotalFOBEmReal_Col))
            
            dTotalCIFMoeda = dTotalCIFMoeda + StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorTotalCIFNaMoeda_Col))
            dTotalCIFReal = dTotalCIFReal + StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorTotalCIFEmReal_Col))
            
            dPesoLiq = dPesoLiq + StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ItemPesoLiq_Col))
            
        End If
        
    Next
    
    LabelTotalFOBAdicaoMoeda.Caption = Format(dTotalFOBMoeda, "standard")
    LabelTotalFOBAdicaoReal.Caption = Format(dTotalFOBReal, "standard")
    LabelTotalCIFAdicaoMoeda.Caption = Format(dTotalCIFMoeda, "standard")
    LabelTotalCIFAdicaoReal.Caption = Format(dTotalCIFReal, "standard")
    LabelPesoLiqAdicao.Caption = Format(dPesoLiq, "standard")

End Sub

Public Sub PC_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_PC_Validate

    If Len(Trim(PC.ClipText)) = 0 Then Exit Sub

    lErro = Long_Critica(PC.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objPedidoCompra.lCodigo = StrParaLong(PC.Text)
    objPedidoCompra.iFilialEmpresa = giFilialEmpresa
    
    'Le o Pedido de Compra
    lErro = CF("PedidoCompra_Le_Numero", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56142 Then gError ERRO_SEM_MENSAGEM
    If lErro = 56142 Then gError 206556

    'Traz para o Frame do pedido de compra
    lErro = Traz_PedidoCompra_Tela(objPedidoCompra)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Exit Sub

Erro_PC_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206556
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", Err, objPedidoCompra.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206557)

    End Select

    Exit Sub

End Sub

Public Sub PC_GotFocus()
Dim iAlt As Integer
     Call MaskEdBox_TrataGotFocus(PC, iAlt)
End Sub

Public Sub PCLabel_Click()

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras
Dim colSelecao As New Collection

On Error GoTo Erro_PCLabel_Click

    objPedidoCompra.lCodigo = StrParaLong(PC.Text)

    'Chama a Tela de browse ("Pedidos de Compra Enviados")
    Call Chama_Tela("PedComprasEnvLista", colSelecao, objPedidoCompra, objEventoPC)

    Exit Sub

Erro_PCLabel_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206558)

    End Select

    Exit Sub

End Sub

Private Sub objEventoPC_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_objEventoPC_evSelecao

    Set objPedidoCompra = obj1

    If Not (objPedidoCompra Is Nothing) Then

        If objPedidoCompra.lNumIntDoc > 0 Then

            'Le o Pedido de Compra
            lErro = CF("PedidoCompras_Le", objPedidoCompra)
            If lErro <> SUCESSO And lErro <> 56118 Then gError ERRO_SEM_MENSAGEM
            If lErro = 56118 Then gError 206559
            
            PC.PromptInclude = False
            PC.Text = objPedidoCompra.lCodigo
            PC.PromptInclude = True

            'Preenche o frame de Pedidos de compra
            lErro = Traz_PedidoCompra_Tela(objPedidoCompra)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        End If

    End If
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoPC_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206559
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", Err, objPedidoCompra.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206560)

    End Select

    Exit Sub

End Sub

Private Function Traz_PedidoCompra_Tela(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Preenche o Frame de Pedido de Compra

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor

On Error GoTo Erro_Traz_PedidoCompra_Tela

    PCData.Caption = Format(objPedidoCompra.dtData, "dd/mm/yyyy")
    PCValor.Caption = Format(objPedidoCompra.dValorTotal, "STANDARD")

    objFornecedor.lCodigo = objPedidoCompra.lFornecedor

    lErro = CF("Fornecedor_Le", objFornecedor)
    If lErro <> SUCESSO And lErro <> 12729 Then gError ERRO_SEM_MENSAGEM
    
    PCForn.Caption = objFornecedor.sNomeReduzido

    objFilialFornecedor.iCodFilial = objPedidoCompra.iFilial

    'Pesquisa se existe filial com o codigo extraido
    lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", objFornecedor.sNomeReduzido, objFilialFornecedor)
    If lErro <> SUCESSO And lErro <> 18272 Then gError ERRO_SEM_MENSAGEM

    'coloca na tela
    PCFilial.Caption = objPedidoCompra.iFilial & SEPARADOR & objFilialFornecedor.sNome

    Traz_PedidoCompra_Tela = SUCESSO
    
    Exit Function

Erro_Traz_PedidoCompra_Tela:

    Traz_PedidoCompra_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206561)

    End Select

    Exit Function
    
End Function

Private Sub BotaoIncluirPC_Click()

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_BotaoIncluirPC_Click

    If Len(Trim(PC.Text)) = 0 Then gError 206562
    If Len(Trim(Moeda1.Text)) = 0 Then gError 206567
    If Len(Trim(TaxaMoeda1.Text)) = 0 Then gError 206568
    
    gbTrazendoPC = True

    objPedidoCompra.lCodigo = StrParaLong(PC.Text)
    objPedidoCompra.iFilialEmpresa = giFilialEmpresa

    lErro = CF("PedidoCompra_Le_Numero", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56142 Then gError ERRO_SEM_MENSAGEM

    'Inclui os dados do pedido de compra no grid de adições e itens
    lErro = Traz_PedidoCompra_DI(objPedidoCompra)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    gbTrazendoPC = False

    Exit Sub

Erro_BotaoIncluirPC_Click:

    gbTrazendoPC = False

    Select Case gErr
    
        Case 206562
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGOPC_NAO_PREENCHIDO", gErr)

        Case 206567
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_NAO_PREENCHIDA", gErr)
            
        Case 206568
            Call Rotina_Erro(vbOKOnly, "ERRO_TAXA_GRID_IMCOMPLETA", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206563)

    End Select

    Exit Sub
    
End Sub

Private Sub Limpa_PC()
'Limpa os dados do Frame de pedidos de Compra

On Error GoTo Erro_Limpa_PC

    PCData.Caption = ""
    PCValor.Caption = ""
    PCForn.Caption = ""
    PCFilial.Caption = ""
    PC.PromptInclude = False
    PC.Text = ""
    PC.PromptInclude = True

    Exit Sub

Erro_Limpa_PC:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206564)

    End Select

    Exit Sub
    
End Sub

'Private Function Traz_PedidoCompra_DI(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'
'Dim lErro As Long
'Dim iLinha As Integer
'Dim objItem As ClassItemPedCompra
'Dim objProd As ClassProduto
'Dim bAchou As Boolean, iAdicao As Integer
'Dim objIPICodigo As ClassClassificacaoFiscal
'Dim sProduto As String, iPreenchido As Integer, iProd As Integer
'Dim dFator As Double, dQtdEst As Double, dQtdAnt As Double, dTaxa As Double, iMoeda As Integer
'Dim colAdicoes As New Collection
'Dim objAdicaoDI As ClassAdicaoDI, colCampos As New Collection, colSaida As New Collection
'
'On Error GoTo Erro_Traz_PedidoCompra_DI
'
'    'Le os itens do Pedido de Compra
'    lErro = CF("ItensPC_Le", objPedidoCompra)
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'    If StrParaInt(MoedaItens.Text) = 1 Then
'        dTaxa = StrParaDbl(TaxaMoeda1.Text)
'        iMoeda = Codigo_Extrai(Moeda1.Text)
'    Else
'        dTaxa = StrParaDbl(TaxaMoeda2.Text)
'        iMoeda = Codigo_Extrai(Moeda2.Text)
'    End If
'
'    'Para cada item
'    For Each objItem In objPedidoCompra.colItens
'
'        Set objProd = New ClassProduto
'
'        objProd.sCodigo = objItem.sProduto
'
'        lErro = CF("Produto_Le", objProd)
'        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
'
'        'Se o produto não tem a classificação fiscal dá erro
'        If Len(Trim(objProd.sIPICodigo)) = 0 Then gError 206565
'
'        'Procura pela classificação fiscal e se não encontrar adiciona, senão não faz nada
'        bAchou = False
''        For iLinha = 1 To objGridAdicao.iLinhasExistentes
''            If Trim(objProd.sIPICodigo) = Trim(GridAdicao.TextMatrix(iLinha, iGrid_IPICodigo_Col)) Then
'        For iLinha = 1 To colAdicoes.Count
'            If Trim(objProd.sIPICodigo) = Trim(colAdicoes.Item(iLinha).sIPICodigo) Then
'                bAchou = True
'                'iAdicao = iLinha
'                colAdicoes.Item(iLinha).colItensAdicaoDI.Add objItem
'                Exit For
'            End If
'        Next
'        If Not bAchou Then
'
''            objGridAdicao.iLinhasExistentes = objGridAdicao.iLinhasExistentes + 1
''            iAdicao = objGridAdicao.iLinhasExistentes
'            Set objIPICodigo = New ClassClassificacaoFiscal
'            Set objAdicaoDI = New ClassAdicaoDI
'
'            objIPICodigo.sCodigo = Trim(objProd.sIPICodigo)
'
'            lErro = CF("ClassificacaoFiscal_Le", objIPICodigo)
'            If lErro <> SUCESSO And lErro <> 123494 Then gError ERRO_SEM_MENSAGEM
'
'            colAdicoes.Add objAdicaoDI
'
'            objAdicaoDI.sIPICodigo = Trim(objProd.sIPICodigo)
'            objAdicaoDI.sDescricao = objIPICodigo.sDescricao
'            objAdicaoDI.dIIAliquota = objIPICodigo.dIIAliquota
'            objAdicaoDI.dPISAliquota = objIPICodigo.dPISAliquota
'            objAdicaoDI.dCOFINSAliquota = objIPICodigo.dCOFINSAliquota
'            objAdicaoDI.dIPIAliquota = objIPICodigo.dIPIAliquota
'            objAdicaoDI.dICMSAliquota = objIPICodigo.dICMSAliquota
'
'            objAdicaoDI.colItensAdicaoDI.Add objItem
'
'
''            'Dados da Classificação Fiscal do Produto
''            GridAdicao.TextMatrix(iAdicao, iGrid_IPICodigo_Col) = Trim(objProd.sIPICodigo)
''            GridAdicao.TextMatrix(iAdicao, iGrid_IPIDescricao_Col) = objIPICodigo.sDescricao
''            GridAdicao.TextMatrix(iAdicao, iGrid_IIAliquota_Col) = Format(objIPICodigo.dIIAliquota, "Percent")
''            GridAdicao.TextMatrix(iAdicao, iGrid_PISAliquota_Col) = Format(objIPICodigo.dPISAliquota, "Percent")
''            GridAdicao.TextMatrix(iAdicao, iGrid_COFINSAliquota_Col) = Format(objIPICodigo.dCOFINSAliquota, "Percent")
''            GridAdicao.TextMatrix(iAdicao, iGrid_IPIAliquota_Col) = Format(objIPICodigo.dIPIAliquota, "Percent")
''            GridAdicao.TextMatrix(iAdicao, iGrid_ICMSAliquota_Col) = Format(objIPICodigo.dICMSAliquota, "Percent")
''
''            'Adiciona a adição na combo do grid de itens
''            Call Carrega_AdicaoItem
'
'        End If
'
'
'    Next
'
'    colCampos.Add "sIPICodigo"
'    Call Ordena_Colecao(colAdicoes, colSaida, colCampos)
'
'    iAdicao = 0
'    For Each objAdicaoDI In colSaida
'
'        iAdicao = iAdicao + 1
'        objGridAdicao.iLinhasExistentes = iAdicao
'
'        'Dados da Classificação Fiscal do Produto
'        GridAdicao.TextMatrix(iAdicao, iGrid_IPICodigo_Col) = Trim(objAdicaoDI.sIPICodigo)
'        GridAdicao.TextMatrix(iAdicao, iGrid_IPIDescricao_Col) = objAdicaoDI.sDescricao
'        GridAdicao.TextMatrix(iAdicao, iGrid_IIAliquota_Col) = Format(objAdicaoDI.dIIAliquota, "Percent")
'        GridAdicao.TextMatrix(iAdicao, iGrid_PISAliquota_Col) = Format(objAdicaoDI.dPISAliquota, "Percent")
'        GridAdicao.TextMatrix(iAdicao, iGrid_COFINSAliquota_Col) = Format(objAdicaoDI.dCOFINSAliquota, "Percent")
'        GridAdicao.TextMatrix(iAdicao, iGrid_IPIAliquota_Col) = Format(objAdicaoDI.dIPIAliquota, "Percent")
'        GridAdicao.TextMatrix(iAdicao, iGrid_ICMSAliquota_Col) = Format(objAdicaoDI.dICMSAliquota, "Percent")
'
'        'Adiciona a adição na combo do grid de itens
'        Call Carrega_AdicaoItem
'
'        'Para cada item
'        For Each objItem In objAdicaoDI.colItensAdicaoDI
'
'            Set objProd = New ClassProduto
'
'            objProd.sCodigo = objItem.sProduto
'
'            lErro = CF("Produto_Le", objProd)
'            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
'
'            'Procura pelo produto\adição e se não encontrar adiciona, senão acerta a quantidade
'            bAchou = False
'            For iLinha = 1 To objGridItens.iLinhasExistentes
'
'                lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProduto, iPreenchido)
'                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'                If iAdicao = StrParaInt(GridItens.TextMatrix(iLinha, iGrid_AdicaoItem_Col)) And sProduto = objProd.sCodigo Then
'                    bAchou = True
'                    iProd = iLinha
'                    Exit For
'                End If
'            Next
'            If Not bAchou Then
'
'                objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
'                iProd = objGridItens.iLinhasExistentes
'
'                lErro = Mascara_RetornaProdutoTela(objProd.sCodigo, sProduto)
'                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'                Call StrParaMasked2(Produto, sProduto)
'    '            Produto.PromptInclude = False
'    '            Produto.Text = sProduto
'    '            Produto.PromptInclude = True
'
'                lErro = CF("UM_Conversao_Trans", objProd.iClasseUM, objItem.sUM, objProd.sSiglaUMEstoque, dFator)
'                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'                'Converte a quantidade do pedido de compra para UM de Estoque
'                dQtdEst = objItem.dQuantidade * dFator
'
'                'Dados do Produto
'                GridItens.TextMatrix(iProd, iGrid_AdicaoItem_Col) = CStr(iAdicao)
'                GridItens.TextMatrix(iProd, iGrid_Produto_Col) = Produto.Text
'                GridItens.TextMatrix(iProd, iGrid_Descricao_Col) = objProd.sDescricao
'                GridItens.TextMatrix(iProd, iGrid_UnidadeMed_Col) = objProd.sSiglaUMEstoque
'                GridItens.TextMatrix(iProd, iGrid_ItemPesoBruto_Col) = Formata_Estoque(objProd.dPesoBruto * dQtdEst)
'                GridItens.TextMatrix(iProd, iGrid_ItemPesoLiq_Col) = Formata_Estoque(objProd.dPesoLiq * dQtdEst)
'
'                'Dados do item do pedido de compra
'                GridItens.TextMatrix(iProd, iGrid_Quantidade_Col) = Formata_Estoque(dQtdEst)
'                If objPedidoCompra.iMoeda = MOEDA_REAL Then
'                    GridItens.TextMatrix(iProd, iGrid_ValorUnitFOBEmReal_Col) = Format(objItem.dPrecoUnitario / dFator, "STANDARD")
'                    GridItens.TextMatrix(iProd, iGrid_ValorUnitFOBNaMoeda_Col) = Format(objItem.dPrecoUnitario / dFator / dTaxa, "STANDARD")
'                ElseIf objPedidoCompra.iMoeda = iMoeda Then
'                    GridItens.TextMatrix(iProd, iGrid_ValorUnitFOBNaMoeda_Col) = Format(objItem.dPrecoUnitario / dFator, "#,##0.00#####")
'                    GridItens.TextMatrix(iProd, iGrid_ValorUnitFOBEmReal_Col) = Format(objItem.dPrecoUnitario / dFator * dTaxa, "#,##0.00#####")
'                End If
'            Else
'
'                lErro = CF("UM_Conversao_Trans", objProd.iClasseUM, objItem.sUM, GridItens.TextMatrix(iProd, iGrid_UnidadeMed_Col), dFator)
'                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'                dQtdAnt = StrParaDbl(GridItens.TextMatrix(iProd, iGrid_Quantidade_Col))
'
'                dQtdEst = dQtdAnt + objItem.dQuantidade * dFator
'
'                GridItens.TextMatrix(iProd, iGrid_Quantidade_Col) = Formata_Estoque(dQtdEst)
'
'                'Acerta o peso de acordo com o que já estava cadastrado
'                If Len(Trim(GridItens.TextMatrix(iProd, iGrid_ItemPesoBruto_Col))) <> 0 Then GridItens.TextMatrix(iProd, iGrid_ItemPesoBruto_Col) = Formata_Estoque(StrParaDbl(GridItens.TextMatrix(iProd, iGrid_ItemPesoBruto_Col)) * (dQtdEst / dQtdAnt))
'                If Len(Trim(GridItens.TextMatrix(iProd, iGrid_ItemPesoLiq_Col))) <> 0 Then GridItens.TextMatrix(iProd, iGrid_ItemPesoLiq_Col) = Formata_Estoque(StrParaDbl(GridItens.TextMatrix(iProd, iGrid_ItemPesoLiq_Col)) * (dQtdEst / dQtdAnt))
'
'            End If
'
'            'Acerta as demais informações
'            Call Calcula_Valores_Adicao(0, 0)
'            Call Calcula_ValorCIF_Linha(iProd)
'            Call Calcula_Valores_Adicao(iAdicao, 0)
'
'        Next
'
'    Next
'
'    'Limpa o Frame de pedido para que possa ser digitado outros
'    Call Limpa_PC
'
'    Traz_PedidoCompra_DI = SUCESSO
'
'    Exit Function
'
'Erro_Traz_PedidoCompra_DI:
'
'    Traz_PedidoCompra_DI = gErr
'
'    Select Case gErr
'
'        Case ERRO_SEM_MENSAGEM
'
'        Case 206565
'            Call Rotina_Erro(vbOKOnly, "ERRO_IPICODIGO_PROD_NAO_PREENCHIDO", gErr, objProd.sCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206566)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Traz_PedidoCompra_DI(ByVal objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long, sProdutoTela As String
Dim iLinha As Integer
Dim objItem As ClassItemPedCompra
Dim objProd As ClassProduto, dFatorVenda As Double
Dim bAchou As Boolean, iAdicao As Integer
Dim objIPICodigo As ClassClassificacaoFiscal
Dim sProduto As String, iPreenchido As Integer, iProd As Integer
Dim dFator As Double, dQtdEst As Double, dQtdAnt As Double, dTaxa As Double, iMoeda As Integer
Dim objAdicaoDI As ClassAdicaoDI, colCampos As New Collection, colSaida As New Collection
Dim objItemAdicaoDI As ClassItemAdicaoDI, objDIInfo As New ClassDIInfo, bAchouItem As Boolean, objItemAdicaoDIAux As ClassItemAdicaoDI

On Error GoTo Erro_Traz_PedidoCompra_DI

    'Le os itens do Pedido de Compra
    lErro = CF("ItensPC_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If StrParaInt(MoedaItens.Text) = 1 Then
        dTaxa = StrParaDbl(TaxaMoeda1.Text)
        iMoeda = Codigo_Extrai(Moeda1.Text)
    Else
        dTaxa = StrParaDbl(TaxaMoeda2.Text)
        iMoeda = Codigo_Extrai(Moeda2.Text)
    End If
    
    lErro = Move_Tela_Memoria(objDIInfo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Para cada item
    For Each objItem In objPedidoCompra.colItens
    
        Set objProd = New ClassProduto
        Set objItemAdicaoDI = New ClassItemAdicaoDI
        
        objProd.sCodigo = objItem.sProduto
    
        lErro = CF("Produto_Le", objProd)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
        
        'Se o produto não tem a classificação fiscal dá erro
        If Len(Trim(objProd.sIPICodigo)) = 0 Then gError 206565
        
        lErro = CF("UM_Conversao_Trans", objProd.iClasseUM, objItem.sUM, objProd.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = CF("UM_Conversao_Trans", objProd.iClasseUM, objItem.sUM, objProd.sSiglaUMVenda, dFatorVenda)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Converte a quantidade do pedido de compra para UM de Estoque
        'dQtdEst = objItem.dQuantidade * dFator
        dQtdEst = (objItem.dQuantidade - objItem.dQuantRecebida - objItem.dQuantRecebimento) * dFator
        
        If dQtdEst > QTDE_ESTOQUE_DELTA Then
            
            objItemAdicaoDI.sProduto = objProd.sCodigo
            objItemAdicaoDI.dPesoBruto = objProd.dPesoBruto * objItem.dQuantidade * dFatorVenda
            objItemAdicaoDI.dPesoLiq = objProd.dPesoLiq * objItem.dQuantidade * dFatorVenda
            objItemAdicaoDI.dQuantidade = dQtdEst
            If objPedidoCompra.iMoeda = MOEDA_REAL Then
                objItemAdicaoDI.dValorUnitFOBEmReal = Format(objItem.dPrecoUnitario / dFator, "STANDARD")
                objItemAdicaoDI.dValorUnitFOBNaMoeda = Format(objItem.dPrecoUnitario / dFator / dTaxa, "STANDARD")
            ElseIf objPedidoCompra.iMoeda = iMoeda Then
                objItemAdicaoDI.dValorUnitFOBNaMoeda = Format(objItem.dPrecoUnitario / dFator, "#,##0.00#####")
                objItemAdicaoDI.dValorUnitFOBEmReal = Format(objItem.dPrecoUnitario / dFator * dTaxa, "#,##0.00#####")
            End If
            objItemAdicaoDI.sUM = objProd.sSiglaUMEstoque
            objItemAdicaoDI.sDescricao = objProd.sDescricao
               
            'incluir o item na aba de pedidos de compra caso ainda nao esteja
            lErro = Mascara_RetornaProdutoTela(objItem.sProduto, sProdutoTela)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            bAchou = False
            For iLinha = 1 To objGridItensPC.iLinhasExistentes
                'se o produto do pedido de compra estiver repetido
                If StrParaLong(GridItensPC.TextMatrix(iLinha, iGrid_CodigoPC_Col)) = objPedidoCompra.lCodigo And GridItensPC.TextMatrix(iLinha, iGrid_ProdutoPC_Col) = sProdutoTela Then
                    bAchou = True
                End If
            Next
            
            If bAchou = False Then
        
                objGridItensPC.iLinhasExistentes = objGridItensPC.iLinhasExistentes + 1
            
                GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes, iGrid_CodigoPC_Col) = objPedidoCompra.lCodigo
                GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes, iGrid_DataPC_Col) = Format(objPedidoCompra.dtData, "dd/mm/yyyy")
                GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes, iGrid_ProdutoPC_Col) = sProdutoTela
                GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes, iGrid_DescProdutoPC_Col) = objItem.sDescProduto
                GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes, iGrid_UMPC_Col) = objItem.sUM
                GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes, iGrid_QuantPC_Col) = Formata_Estoque(objItem.dQuantidade - objItem.dQuantRecebida - objItem.dQuantRecebimento)
                
            End If
            
            'Procura pela classificação fiscal e se não encontrar adiciona, senão não faz nada
            bAchou = False
            For iLinha = 1 To objDIInfo.colAdicoesDI.Count
                If Trim(objProd.sIPICodigo) = Trim(objDIInfo.colAdicoesDI.Item(iLinha).sIPICodigo) Then
                    bAchou = True
                    bAchouItem = False
                    For Each objItemAdicaoDIAux In objDIInfo.colAdicoesDI.Item(iLinha).colItensAdicaoDI
                        If objItemAdicaoDIAux.sProduto = objItemAdicaoDI.sProduto Then
                            objItemAdicaoDIAux.dQuantidade = objItemAdicaoDIAux.dQuantidade + objItemAdicaoDI.dQuantidade
                            objItemAdicaoDIAux.dPesoBruto = objItemAdicaoDIAux.dPesoBruto + objItemAdicaoDI.dPesoBruto
                            objItemAdicaoDIAux.dPesoLiq = objItemAdicaoDIAux.dPesoLiq + objItemAdicaoDI.dPesoLiq
                            bAchouItem = True
                            Exit For
                        End If
                    Next
                    If Not bAchouItem Then objDIInfo.colAdicoesDI.Item(iLinha).colItensAdicaoDI.Add objItemAdicaoDI
                    Exit For
                End If
            Next
            If Not bAchou Then
            
                Set objIPICodigo = New ClassClassificacaoFiscal
                Set objAdicaoDI = New ClassAdicaoDI
            
                objIPICodigo.sCodigo = Trim(objProd.sIPICodigo)
                
                lErro = CF("ClassificacaoFiscal_Le", objIPICodigo)
                If lErro <> SUCESSO And lErro <> 123494 Then gError ERRO_SEM_MENSAGEM
                
                objDIInfo.colAdicoesDI.Add objAdicaoDI
                
                objAdicaoDI.sIPICodigo = Trim(objProd.sIPICodigo)
                objAdicaoDI.dIIAliquota = objIPICodigo.dIIAliquota
                objAdicaoDI.dPISAliquota = objIPICodigo.dPISAliquota
                objAdicaoDI.dCOFINSAliquota = objIPICodigo.dCOFINSAliquota
                objAdicaoDI.dIPIAliquota = objIPICodigo.dIPIAliquota
                objAdicaoDI.dICMSAliquota = objIPICodigo.dICMSAliquota
                
                objAdicaoDI.colItensAdicaoDI.Add objItemAdicaoDI
                       
            End If
            
        End If
        
    Next
       
    colCampos.Add "sIPICodigo"
    Call Ordena_Colecao(objDIInfo.colAdicoesDI, colSaida, colCampos)
    
    iAdicao = 0
    For Each objAdicaoDI In colSaida
        iAdicao = iAdicao + 1
        objAdicaoDI.iSeq = iAdicao
        For Each objItemAdicaoDI In objAdicaoDI.colItensAdicaoDI
            objItemAdicaoDI.iAdicao = objAdicaoDI.iSeq
        Next
    Next
    
    Set objDIInfo.colAdicoesDI = colSaida
   
    lErro = Preenche_GridAdicao_Tela(objDIInfo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
       
    iAdicao = 0
    iProd = 0
    For Each objAdicaoDI In objDIInfo.colAdicoesDI
        
        iAdicao = iAdicao + 1
        
        For Each objItemAdicaoDI In objAdicaoDI.colItensAdicaoDI
        
            iProd = iProd + 1
        
            'Acerta as demais informações
            Call Calcula_Valores_Adicao(0, 0)
            Call Calcula_ValorCIF_Linha(iProd)
            Call Calcula_Valores_Adicao(iAdicao, 0)
            
        Next
    
    Next

    'Limpa o Frame de pedido para que possa ser digitado outros
    Call Limpa_PC

    Traz_PedidoCompra_DI = SUCESSO
    
    Exit Function

Erro_Traz_PedidoCompra_DI:

    Traz_PedidoCompra_DI = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 206565
            Call Rotina_Erro(vbOKOnly, "ERRO_IPICODIGO_PROD_NAO_PREENCHIDO", gErr, objProd.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206566)

    End Select

    Exit Function
    
End Function

Private Sub UpDownDData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDData_DownClick

    DData.SetFocus

    If Len(DData.ClipText) > 0 Then

        sData = DData.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 196603

        DData.Text = sData

    End If

    Exit Sub

Erro_UpDownDData_DownClick:

    Select Case gErr

        Case 196603

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196604)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDData_UpClick

    DData.SetFocus

    If Len(Trim(DData.ClipText)) > 0 Then

        sData = DData.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 196605

        DData.Text = sData

    End If

    Exit Sub

Erro_UpDownDData_UpClick:

    Select Case gErr

        Case 196605

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196606)

    End Select

    Exit Sub

End Sub

Private Sub DData_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DData, iAlterado)
    
End Sub

Private Sub DData_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DData_Validate

    If Len(Trim(DData.ClipText)) <> 0 Then

        lErro = Data_Critica(DData.Text)
        If lErro <> SUCESSO Then gError 196607

    End If

    Exit Sub

Erro_DData_Validate:

    Cancel = True

    Select Case gErr

        Case 196607

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196608)

    End Select

    Exit Sub

End Sub

Private Sub DData_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DUF_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DLocal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub BotaoTrazCotacao1_Click()
'Traz a última cotação da Moeda1 selecionada

Dim lErro As Long
Dim objCotacao As New ClassCotacaoMoeda
Dim objCotacaoAnterior As New ClassCotacaoMoeda

On Error GoTo Erro_BotaoTrazCotacao1_Click

    'Carrega objCotacao
    objCotacao.dtData = gdtDataAtual
    
    'Se a Moeda1 não foi selecionada => Erro
    If Len(Trim(Moeda1.Text)) = 0 Then gError 196657
        
    'Preeche com a Moeda1 selecionada
    objCotacao.iMoeda = Codigo_Extrai(Moeda1.List(Moeda1.ListIndex))
    objCotacaoAnterior.iMoeda = Codigo_Extrai(Moeda1.List(Moeda1.ListIndex))

    'Chama função de leitura
    lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao, objCotacaoAnterior)
    If lErro <> SUCESSO Then gError 196658
    
    'Se nao existe Cotacao1 para a data informada => Mostra a última.
    TaxaMoeda1.Text = IIf(objCotacao.dValor <> 0, Format(objCotacao.dValor, TaxaMoeda1.Format), Format(objCotacaoAnterior.dValor, TaxaMoeda1.Format))
    
    Call ComparativoMoedaReal_Calcula(1, StrParaDbl(TaxaMoeda1.Text))

    Exit Sub
    
Erro_BotaoTrazCotacao1_Click:

    Select Case gErr
    
        Case 196657
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_NAO_PREENCHIDA", gErr)
            '??? Falta cadastrar: ERRO_Moeda1_NAO_PREENCHIDA - "Para trazer a cotação a Moeda1 deve ser selecionada antes."
            
        Case 196658
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 196659)
    
    End Select
    
End Sub

Public Sub Moeda1_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TaxaMoeda1_Change()
    iAlterado = REGISTRO_ALTERADO
    iTaxaMoeda1Alterada = REGISTRO_ALTERADO
End Sub

Public Sub TaxaMoeda1_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TaxaMoeda1_Validate

    If iTaxaMoeda1Alterada <> 0 Then
    
        'Verifica se algum valor foi digitado
        If Len(Trim(TaxaMoeda1.Text)) > 0 Then
    
            'Critica se é valor Positivo
            lErro = Valor_Positivo_Critica_Double(TaxaMoeda1.Text)
            If lErro <> SUCESSO Then gError 196660
        
            'Põe o valor formatado na tela
            TaxaMoeda1.Text = Format(TaxaMoeda1.Text, TaxaMoeda1.Format)
            
            'Calcula o comparativo em real para o grid de itens
            Call ComparativoMoedaReal_Calcula(1, StrParaDbl(TaxaMoeda1.Text))
        
        End If
    
        iTaxaMoeda1Alterada = 0
    
    End If
    
    Exit Sub

Erro_TaxaMoeda1_Validate:

    Cancel = True

    Select Case gErr

        Case 196660

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196661)

    End Select

    Exit Sub

End Sub

Public Sub BotaoTrazCotacao2_Click()
'Traz a última cotação da Moeda2 selecionada

Dim lErro As Long
Dim objCotacao As New ClassCotacaoMoeda
Dim objCotacaoAnterior As New ClassCotacaoMoeda

On Error GoTo Erro_BotaoTrazCotacao2_Click

    'Carrega objCotacao
    objCotacao.dtData = gdtDataAtual
    
    'Se a Moeda2 não foi selecionada => Erro
    If Len(Trim(Moeda2.Text)) = 0 Then gError 196657
        
    'Preeche com a Moeda2 selecionada
    objCotacao.iMoeda = Codigo_Extrai(Moeda2.List(Moeda2.ListIndex))
    objCotacaoAnterior.iMoeda = Codigo_Extrai(Moeda2.List(Moeda2.ListIndex))

    'Chama função de leitura
    lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao, objCotacaoAnterior)
    If lErro <> SUCESSO Then gError 196658
    
    'Se nao existe Cotacao2 para a data informada => Mostra a última.
    TaxaMoeda2.Text = IIf(objCotacao.dValor <> 0, Format(objCotacao.dValor, TaxaMoeda2.Format), Format(objCotacaoAnterior.dValor, TaxaMoeda2.Format))
    
    Call ComparativoMoedaReal_Calcula(2, StrParaDbl(TaxaMoeda2.Text))

    Exit Sub
    
Erro_BotaoTrazCotacao2_Click:

    Select Case gErr
    
        Case 196657
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_NAO_PREENCHIDA", gErr)
            '??? Falta cadastrar: ERRO_Moeda2_NAO_PREENCHIDA - "Para trazer a cotação a Moeda2 deve ser selecionada antes."
            
        Case 196658
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 196659)
    
    End Select
    
End Sub

Public Sub Moeda2_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TaxaMoeda2_Change()
    iAlterado = REGISTRO_ALTERADO
    iTaxaMoeda2Alterada = REGISTRO_ALTERADO
End Sub

Public Sub TaxaMoeda2_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TaxaMoeda2_Validate

    If iTaxaMoeda2Alterada <> 0 Then
    
        'Verifica se algum valor foi digitado
        If Len(Trim(TaxaMoeda2.Text)) > 0 Then
    
            'Critica se é valor Positivo
            lErro = Valor_Positivo_Critica_Double(TaxaMoeda2.Text)
            If lErro <> SUCESSO Then gError 196660
        
            'Põe o valor formatado na tela
            TaxaMoeda2.Text = Format(TaxaMoeda2.Text, TaxaMoeda2.Format)
            
            'Calcula o comparativo em real para o grid de itens
            Call ComparativoMoedaReal_Calcula(2, StrParaDbl(TaxaMoeda2.Text))
        
        End If
    
        iTaxaMoeda2Alterada = 0
    
    End If
    
    Exit Sub

Erro_TaxaMoeda2_Validate:

    Cancel = True

    Select Case gErr

        Case 196660

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196661)

    End Select

    Exit Sub

End Sub

Private Sub MoedaMercadoria_Change()
    Call ValorMercadoriaMoeda_Validate(bSGECancelDummy)
End Sub

Private Sub MoedaFrete_Change()
    Call ValorFreteInternacMoeda_Validate(bSGECancelDummy)
End Sub

Private Sub MoedaSeguro_Change()
    Call ValorSeguroInternacMoeda_Validate(bSGECancelDummy)
End Sub

Private Sub MoedaMercadoria_Click()
    Call ValorMercadoriaMoeda_Validate(bSGECancelDummy)
End Sub

Private Sub MoedaFrete_Click()
    Call ValorFreteInternacMoeda_Validate(bSGECancelDummy)
End Sub

Private Sub MoedaSeguro_Click()
    Call ValorSeguroInternacMoeda_Validate(bSGECancelDummy)
End Sub

Private Sub MoedaItens_Change()
    If StrParaInt(MoedaItens.Text) <> iMoedaItensAnt Then
        iMoedaItensAnt = StrParaInt(MoedaItens.Text)
        Call RecalcularCIFs
    End If
End Sub

Private Sub MoedaItens_Click()
    If StrParaInt(MoedaItens.Text) <> iMoedaItensAnt Then
        iMoedaItensAnt = StrParaInt(MoedaItens.Text)
        Call RecalcularCIFs
    End If
End Sub

Private Sub CodFabricante_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodFabricante_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub CodFabricante_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub CodFabricante_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = CodFabricante
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TaxaSisComex_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TaxaSisComex_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub TaxaSisComex_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub TaxaSisComex_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = TaxaSiscomex
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DespAdua_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DespAdua_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub DespAdua_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub DespAdua_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = DespAdua
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ICMSPercRedBase_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ICMSPercRedBase_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub ICMSPercRedBase_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub ICMSPercRedBase_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = ICMSPercRedBase
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Inicializa_GridItensPC(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens PC

    Set objGridInt = New AdmGrid

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Pedido Compra")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (CodigoPC.Name)
    objGridInt.colCampo.Add (DataPC.Name)
    objGridInt.colCampo.Add (ProdutoPC.Name)
    objGridInt.colCampo.Add (DescProdutoPC.Name)
    objGridInt.colCampo.Add (UMPC.Name)
    objGridInt.colCampo.Add (QuantPC.Name)

    iGrid_CodigoPC_Col = 1
    iGrid_DataPC_Col = 2
    iGrid_ProdutoPC_Col = 3
    iGrid_DescProdutoPC_Col = 4
    iGrid_UMPC_Col = 5
    iGrid_QuantPC_Col = 6


    'Grid do GridInterno
    objGridInt.objGrid = GridItensPC

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_PEDIDO_COMPRAS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 18

    'Largura da primeira coluna
    GridItensPC.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridItensPC = SUCESSO

    Exit Function

End Function

Function Saida_Celula_GridItensPC(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItensPC

    'Verifica qual a coluna do Grid em questão
    Select Case objGridInt.objGrid.Col

        Case iGrid_CodigoPC_Col
            lErro = Saida_Celula_CodigoPC(objGridInt)
            If lErro <> SUCESSO Then gError 210568

        Case iGrid_ProdutoPC_Col
            lErro = Saida_Celula_ProdutoPC(objGridInt)
            If lErro <> SUCESSO Then gError 210569

        'Recebido
        Case iGrid_QuantPC_Col
            lErro = Saida_Celula_QuantPC(objGridInt)
            If lErro <> SUCESSO Then gError 210570

    End Select

    Saida_Celula_GridItensPC = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItensPC:

    Saida_Celula_GridItensPC = gErr

    Select Case gErr

        Case 210568 To 210570

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210571)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CodigoPC(objGridInt As AdmGrid) As Long
'faz a critica da celula de codigo PC do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_CodigoPC

    Set objGridInt.objControle = CodigoPC
    
    If Len(Trim(CodigoPC.Text)) > 0 Then
    
        lErro = Long_Critica(CodigoPC.Text)
        If lErro <> SUCESSO Then gError 210573
        
        objPedidoCompra.lCodigo = StrParaLong(CodigoPC.Text)
        objPedidoCompra.iFilialEmpresa = giFilialEmpresa
        
        'Busca no BD Pedido de Compra com FilialEmpresa e Codigo passados como parametros
      
        lErro = CF("PedidoCompra_Le_Numero", objPedidoCompra)
        If lErro <> SUCESSO And lErro <> 56142 Then gError 210574
        
        If lErro = 56142 Then gError 210575
        
        If objPedidoCompra.dtDataEnvio = DATA_NULA Then gError 216879
        
        'verifica se precisa preencher o grid com uma nova linha
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
        GridItensPC.TextMatrix(objGridInt.objGrid.Row, iGrid_DataPC_Col) = Format(objPedidoCompra.dtData, "dd/mm/yyyy")
        
    Else
        
        GridItensPC.TextMatrix(objGridInt.objGrid.Row, iGrid_DataPC_Col) = ""
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 210576


    Saida_Celula_CodigoPC = SUCESSO

    Exit Function

Erro_Saida_Celula_CodigoPC:

    Saida_Celula_CodigoPC = gErr

    Select Case gErr

        Case 210573, 210574, 210576
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 210575
            Call Rotina_Erro(vbOKOnly, "ERRO_PC_BAIXADO_OU_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 216879
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_ENVIADO", gErr, objPedidoCompra.lCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210577)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ProdutoPC(objGridInt As AdmGrid) As Long
'faz a critica da celula de item PC do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String
Dim iLinha As Integer
Dim sProdutoFormatado As String
Dim iIndice As Integer
Dim objItemPC As ClassItemPedCompra
Dim iAchou As Integer

On Error GoTo Erro_Saida_Celula_ProdutoPC

    Set objGridInt.objControle = ProdutoPC
    
    'Verifica se o Produto está preenchido
    If Len(Trim(ProdutoPC.ClipText)) > 0 Then

        lErro = CF("Produto_Formata", ProdutoPC.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 210579

        objProduto.sCodigo = sProdutoFormatado

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            objPedidoCompra.lCodigo = StrParaLong(GridItensPC.TextMatrix(GridItensPC.Row, iGrid_CodigoPC_Col))
            objPedidoCompra.iFilialEmpresa = giFilialEmpresa

            'Lê os itens de um Pedido de Compra a partir do código e de FilialEmpresa do Pedido de Compras
            lErro = CF("ItensPC_Le_Codigo", objPedidoCompra)
            If lErro <> SUCESSO And lErro <> 25605 Then gError 210580

            iAchou = 0

            For Each objItemPC In objPedidoCompra.colItens

                If objItemPC.sProduto = sProdutoFormatado Then
                    iAchou = 1
                    Exit For
                End If
                
            Next
            
            If iAchou = 0 Then gError 210581
            
            For iIndice = 1 To objGridInt.iLinhasExistentes
                If iIndice <> GridItensPC.Row Then
                    'se o produto do pedido de compra estiver repetido
                    If GridItensPC.TextMatrix(iIndice, iGrid_CodigoPC_Col) = GridItensPC.TextMatrix(GridItensPC.Row, iGrid_CodigoPC_Col) And _
                    GridItensPC.TextMatrix(iIndice, iGrid_ProdutoPC_Col) = ProdutoPC.Text Then 'GridItensPC.TextMatrix(GridItensPC.Row, iGrid_ProdutoPC_Col) Then
                        gError 210585
                    End If
                End If
            Next
            
            lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
            If lErro <> SUCESSO Then gError 210582

            'Coloca o produto no grid
            ProdutoPC.PromptInclude = False
            ProdutoPC.Text = sProduto
            ProdutoPC.PromptInclude = True


        End If

        GridItensPC.TextMatrix(GridItensPC.Row, iGrid_DescProdutoPC_Col) = objItemPC.sDescProduto
        GridItensPC.TextMatrix(GridItensPC.Row, iGrid_UMPC_Col) = objItemPC.sUM

    Else
        GridItensPC.TextMatrix(GridItensPC.Row, iGrid_DescProdutoPC_Col) = ""
        GridItensPC.TextMatrix(GridItensPC.Row, iGrid_UMPC_Col) = ""
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 210583

    Saida_Celula_ProdutoPC = SUCESSO

    Exit Function

Erro_Saida_Celula_ProdutoPC:

    Saida_Celula_ProdutoPC = gErr

    Select Case gErr

        Case 210579, 210580, 210582, 210583
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 210581
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO_PC", gErr, sProdutoFormatado)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 210585
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_CADASTRADO_GRIDITENSPC", gErr, sProdutoFormatado, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210584)

    End Select

    Exit Function

End Function


Private Function Saida_Celula_QuantPC(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double
Dim iAchou As Integer
Dim objPedidoCompra As New ClassPedidoCompras
Dim objProduto As New ClassProduto
Dim objItemPC As ClassItemPedCompra
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim sProdutoFormatado As String

On Error GoTo Erro_Saida_Celula_QuantPC

    Set objGridInt.objControle = QuantPC
    QuantPC.Text = Trim(QuantPC.Text)
    
    'Se quantidade estiver preenchida
    If Len(Trim(QuantPC.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(QuantPC.Text)
        If lErro <> SUCESSO Then gError 210586

        dQuantidade = StrParaDbl(QuantPC.Text)

        'Formata o produto
        lErro = CF("Produto_Formata", GridItensPC.TextMatrix(GridItensPC.Row, iGrid_ProdutoPC_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 210587
        


        objPedidoCompra.lCodigo = StrParaLong(GridItensPC.TextMatrix(GridItensPC.Row, iGrid_CodigoPC_Col))
        objPedidoCompra.iFilialEmpresa = giFilialEmpresa

        'Lê os itens de um Pedido de Compra a partir do código e de FilialEmpresa do Pedido de Compras
        lErro = CF("ItensPC_Le_Codigo", objPedidoCompra)
        If lErro <> SUCESSO And lErro <> 25605 Then gError 210588

        iAchou = 0

        For Each objItemPC In objPedidoCompra.colItens

            If objItemPC.sProduto = sProdutoFormatado Then
                iAchou = 1
                Exit For
            End If
                
        Next
            
        If iAchou = 0 Then gError 210589
        
        If dQuantidade > (objItemPC.dQuantidade - objItemPC.dQuantRecebida - objItemPC.dQuantRecebimento) Then
            gError 210590
        End If
        

        'Coloca o valor Formatado na tela
        QuantPC.Text = Formata_Estoque(dQuantidade)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 210591

    Saida_Celula_QuantPC = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantPC:

    Saida_Celula_QuantPC = gErr

    Select Case gErr

        Case 210586 To 210588
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 210589
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO_PC", gErr, sProdutoFormatado)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 210590
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_MAIOR_NAORECEBIDA_PC", gErr, sProdutoFormatado, dQuantidade, (objItemPC.dQuantidade - objItemPC.dQuantRecebida - objItemPC.dQuantRecebimento))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210592)

    End Select

    Exit Function

End Function

Private Sub GridItensPC_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItensPC, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensPC, iAlterado)
    End If

End Sub

Private Sub GridItensPC_GotFocus()
    Call Grid_Recebe_Foco(objGridItensPC)
End Sub

Private Sub GridItensPC_EnterCell()
    Call Grid_Entrada_Celula(objGridItensPC, iAlterado)
End Sub

Private Sub GridItensPC_LeaveCell()
    Call Saida_Celula(objGridItensPC)
End Sub

Private Sub GridItensPC_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItensPC, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensPC, iAlterado)
    End If

End Sub

Private Sub GridItensPC_RowColChange()
    Call Grid_RowColChange(objGridItensPC)
    If GridItensPC.Row <> 0 Then Call Atualiza_Totais_Adicao(GridItensPC.Row)
End Sub

Private Sub GridItensPC_Scroll()
    Call Grid_Scroll(objGridItensPC)
End Sub

Private Sub GridItensPC_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer

On Error GoTo Erro_GridItensPC_KeyDown

    iLinhasExistentesAnterior = objGridItensPC.iLinhasExistentes
    iItemAtual = GridItensPC.Row

    Call Grid_Trata_Tecla1(KeyCode, objGridItensPC)
    
    Exit Sub
    
Erro_GridItensPC_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210595)

    End Select

    Exit Sub
    
End Sub

Public Sub GridItensPC_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItensPC)
    
End Sub

Public Sub ProdutoPC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ProdutoPC_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensPC)

End Sub

Public Sub ProdutoPC_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensPC)

End Sub

Public Sub ProdutoPC_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensPC.objControle = ProdutoPC
    lErro = Grid_Campo_Libera_Foco(objGridItensPC)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub CodigoPC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CodigoPC_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensPC)

End Sub

Public Sub CodigoPC_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensPC)

End Sub

Public Sub CodigoPC_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensPC.objControle = CodigoPC
    lErro = Grid_Campo_Libera_Foco(objGridItensPC)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub QuantPC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantPC_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensPC)

End Sub

Public Sub QuantPC_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensPC)

End Sub

Public Sub QuantPC_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensPC.objControle = QuantPC
    lErro = Grid_Campo_Libera_Foco(objGridItensPC)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function Move_GridItensPC_Memoria(objDIInfo As ClassDIInfo) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objAdicaoDI As ClassAdicaoDI
Dim objItemAdicaoDI As ClassItemAdicaoDI
Dim bAchou As Boolean
Dim iItem As Integer
Dim sProduto As String
Dim iPreenchido As Integer
Dim objItemPCDI As ClassItemPCDI
Dim iAchou As Integer

On Error GoTo Erro_Move_GridItensPC_Memoria

    For iLinha = 1 To objGridItensPC.iLinhasExistentes
    
        
        Set objItemPCDI = New ClassItemPCDI
    
        'Formata o produto
        lErro = CF("Produto_Formata", GridItensPC.TextMatrix(iLinha, iGrid_ProdutoPC_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 210594

        objItemPCDI.sProdutoPC = sProduto
        
        bAchou = False
        For Each objAdicaoDI In objDIInfo.colAdicoesDI
            For Each objItemAdicaoDI In objAdicaoDI.colItensAdicaoDI
                If objItemAdicaoDI.sProduto = objItemPCDI.sProdutoPC Then
                    bAchou = True
                    Exit For
                End If
            Next
            If iAchou = True Then Exit For
        Next
        
        If Not bAchou Then gError 210596
    
        objItemPCDI.lCodigoPC = StrParaLong(GridItensPC.TextMatrix(iLinha, iGrid_CodigoPC_Col))
        objItemPCDI.dtDataPC = StrParaDate(GridItensPC.TextMatrix(iLinha, iGrid_DataPC_Col))
        objItemPCDI.sDescProdPC = GridItensPC.TextMatrix(iLinha, iGrid_DescProdutoPC_Col)
        objItemPCDI.sUMPC = GridItensPC.TextMatrix(iLinha, iGrid_UMPC_Col)
        objItemPCDI.iSeq = iLinha
        objItemPCDI.iFilialEmpresa = giFilialEmpresa
        objItemPCDI.dQuantPC = StrParaDbl(GridItensPC.TextMatrix(iLinha, iGrid_QuantPC_Col))
        
        objDIInfo.colItensPC.Add objItemPCDI
    
    Next

    Move_GridItensPC_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItensPC_Memoria:

    Move_GridItensPC_Memoria = gErr

    Select Case gErr
    
        Case 210594
        
        Case 210596
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOPC_DI_INEXISTENTE", gErr, sProduto, iLinha)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210597)

    End Select

    Exit Function

End Function

Function Preenche_GridItensPCDI_Tela(objDIInfo As ClassDIInfo) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objItemPCDI As ClassItemPCDI
Dim sProdutoEnxuto As String

On Error GoTo Erro_Preenche_GridItensPCDI_Tela

    Call Grid_Limpa(objGridItensPC)

    iLinha = 0
    
    For Each objItemPCDI In objDIInfo.colItensPC
    
            iLinha = iLinha + 1
            
            lErro = Mascara_RetornaProdutoEnxuto(objItemPCDI.sProdutoPC, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError 210606

            'Call StrParaMasked2(Produto, sProdutoEnxuto)
            ProdutoPC.PromptInclude = False
            ProdutoPC.Text = sProdutoEnxuto
            ProdutoPC.PromptInclude = True
            
    
            GridItensPC.TextMatrix(iLinha, iGrid_CodigoPC_Col) = objItemPCDI.lCodigoPC
            GridItensPC.TextMatrix(iLinha, iGrid_DataPC_Col) = Format(objItemPCDI.dtDataPC, "dd/mm/yyyy")
   
            GridItensPC.TextMatrix(iLinha, iGrid_ProdutoPC_Col) = ProdutoPC.Text
            GridItensPC.TextMatrix(iLinha, iGrid_DescProdutoPC_Col) = objItemPCDI.sDescProdPC
            GridItensPC.TextMatrix(iLinha, iGrid_UMPC_Col) = objItemPCDI.sUMPC
            GridItensPC.TextMatrix(iLinha, iGrid_QuantPC_Col) = Formata_Estoque(objItemPCDI.dQuantPC)
            

    Next

    objGridItensPC.iLinhasExistentes = iLinha
    Call Grid_Refresh_Checkbox(objGridItensPC)
    
    Preenche_GridItensPCDI_Tela = SUCESSO

    Exit Function

Erro_Preenche_GridItensPCDI_Tela:

    Preenche_GridItensPCDI_Tela = gErr

    Select Case gErr
    
        Case 210606

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210607)

    End Select

    Exit Function

End Function

Private Sub BotaoItensPC_Click()
       
Dim objItemPC As New ClassItensPedCompraPC
Dim lErro As Long
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_BotaoItensPC_Click

    If GridItensPC.Row = 0 Then gError 210622

    If Me.ActiveControl Is CodigoPC Then
        objItemPC.lCodigo = StrParaLong(CodigoPC.Text)
    Else
        
        objItemPC.lCodigo = StrParaLong(GridItensPC.TextMatrix(GridItensPC.Row, iGrid_CodigoPC_Col))
    End If

    colSelecao.Add DATA_NULA
    sSelecao = "StatusBaixa = 0 AND DataEnvio <> ? "
    
    Call Chama_Tela("ItensPedCompraPCLista", colSelecao, objItemPC, objEventoItemPC, sSelecao)

    Exit Sub
    
Erro_BotaoItensPC_Click:

    Select Case gErr

        Case 210622
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210623)

    End Select

    Exit Sub

End Sub

Private Sub objEventoItemPC_evSelecao(obj1 As Object)

Dim objItemPC As ClassItensPedCompraPC
Dim lErro As Long
Dim sProdutoMascarado As String
Dim iLinha As Integer
Dim iIndice As Integer

On Error GoTo Erro_objEventoItemPC_evSelecao

    Set objItemPC = obj1

    lErro = Mascara_RetornaProdutoEnxuto(objItemPC.sProduto, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 210614

    ProdutoPC.PromptInclude = False
    ProdutoPC.Text = sProdutoMascarado
    ProdutoPC.PromptInclude = True

    For iIndice = 1 To objGridItensPC.iLinhasExistentes
        If iIndice <> GridItensPC.Row Then
            'se o produto do pedido de compra estiver repetido
            If StrParaLong(GridItensPC.TextMatrix(iIndice, iGrid_CodigoPC_Col)) = objItemPC.lCodigo And _
            GridItensPC.TextMatrix(iIndice, iGrid_ProdutoPC_Col) = ProdutoPC.Text Then
                gError 210624
            End If
        End If
    Next



    GridItensPC.TextMatrix(GridItensPC.Row, iGrid_CodigoPC_Col) = objItemPC.lCodigo
    CodigoPC.Text = objItemPC.lCodigo
        
    GridItensPC.TextMatrix(GridItensPC.Row, iGrid_DataPC_Col) = Format(objItemPC.dtData, "dd/mm/yyyy")
    GridItensPC.TextMatrix(GridItensPC.Row, iGrid_ProdutoPC_Col) = ProdutoPC.Text
    GridItensPC.TextMatrix(GridItensPC.Row, iGrid_DescProdutoPC_Col) = objItemPC.sDescProduto
    GridItensPC.TextMatrix(GridItensPC.Row, iGrid_UMPC_Col) = objItemPC.sUM

    If (GridItensPC.Row - GridItensPC.FixedRows) = objGridItensPC.iLinhasExistentes Then
        objGridItensPC.iLinhasExistentes = objGridItensPC.iLinhasExistentes + 1
    End If


    Me.Show

    Exit Sub

Erro_objEventoItemPC_evSelecao:

    Select Case gErr

        Case 210614
        
        Case 210624
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_CADASTRADO_GRIDITENSPC", gErr, sProdutoMascarado, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210615)

    End Select

    Exit Sub

End Sub

Private Sub Calcula_Siscomex_Adicao(ByVal iAdicao As Integer, ByVal dTaxaSiscomexDI As Double, dTaxaSiscomex As Double)
'http://socomex.blogspot.com.br/2011/05/siscomex-taxa-de-utilizacao-novos.html

Dim iNumAdicoes As Integer, dAcumulado As Double, iLinha As Integer

    iNumAdicoes = objGridAdicao.iLinhasExistentes

    If iNumAdicoes = 1 Then
    
        dTaxaSiscomex = dTaxaSiscomexDI
    
    Else
    
        If iAdicao <> iNumAdicoes Then
        
            dTaxaSiscomex = 185 / iNumAdicoes
            
            Select Case iAdicao
            
                Case 1, 2
                    dTaxaSiscomex = dTaxaSiscomex + 29.5
                
                Case 3, 4, 5
                    dTaxaSiscomex = dTaxaSiscomex + 23.6
                
                Case 6 To 10
                    dTaxaSiscomex = dTaxaSiscomex + 17.7
                    
                Case 11 To 20
                    dTaxaSiscomex = dTaxaSiscomex + 11.8
                    
                Case 21 To 50
                    dTaxaSiscomex = dTaxaSiscomex + 5.9
                    
                Case Else
                    dTaxaSiscomex = dTaxaSiscomex + 2.95
                
            End Select
            
        Else
        
            For iLinha = 1 To iNumAdicoes - 1
            
                dAcumulado = dAcumulado + StrParaDbl(GridAdicao.TextMatrix(iLinha, iGrid_TaxaSiscomex_Col))
                
            Next
            
            dTaxaSiscomex = Arredonda_Moeda(dTaxaSiscomexDI - dAcumulado, 2)
            
        End If
    
    End If

End Sub

'-------------------------------------NFe 3.10------------------------------------------
Private Sub CNPJAdquir_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CNPJAdquir_GotFocus()
    Call MaskEdBox_TrataGotFocus(CNPJAdquir, iAlterado)
End Sub

Private Sub CNPJAdquir_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CNPJAdquir_Validate

    If Len(Trim(CNPJAdquir.Text)) = 0 Then Exit Sub
    
    'Pelo Tamanho verifica se é CPF ou CGC
    Select Case Len(Trim(CNPJAdquir.Text))
                
        Case STRING_CGC  'CGC

            lErro = Cgc_Critica(CNPJAdquir.Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            CNPJAdquir.Format = "00\.000\.000\/0000-00; ; ; "
            CNPJAdquir.Text = CNPJAdquir.Text
            
        Case Else

            gError 213585

    End Select

    Exit Sub

Erro_CNPJAdquir_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 213585
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213586)

    End Select

    Exit Sub

End Sub

Private Sub Intermedio_Change()
    iAlterado = REGISTRO_ALTERADO
    Call Trata_Intermedio
End Sub

Private Sub Intermedio_Click()
    Call Trata_Intermedio
End Sub

Sub Trata_Intermedio()
'--1=Importação por conta própria;
'--2=Importação por conta e ordem;
'--3=Importação por encomenda;
    If Codigo_Extrai(Intermedio.Text) = 1 Then
        FrameAdquirente.Enabled = False
        CNPJAdquir.Text = ""
        UFAdquir.ListIndex = -1
    Else
        FrameAdquirente.Enabled = True
    End If
End Sub

Private Sub NumDrawback_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumDrawback_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdicao)
End Sub

Private Sub NumDrawback_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdicao)
End Sub

Private Sub NumDrawback_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdicao.objControle = NumDrawback
    lErro = Grid_Campo_Libera_Foco(objGridAdicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub
'-------------------------------------NFe 3.10------------------------------------------

Function Traz_DIInfo_Tela_XML(objDIInfo As ClassDIInfo) As Long

Dim lErro As Long, iIndice As Integer

On Error GoTo Erro_Traz_DIInfo_Tela_XML
   
    Call DateParaMasked(Data, objDIInfo.dtData)
    gdtDataAnterior = objDIInfo.dtData

    DIDescricao.Text = objDIInfo.sDescricao
    
    For iIndice = 0 To Moeda1.ListCount - 1
        If objDIInfo.iMoeda1 = Codigo_Extrai(Moeda1.List(iIndice)) Then
            Moeda1.ListIndex = iIndice
            Exit For
        End If
    Next

    If objDIInfo.dTaxaMoeda1 <> 0 Then
        TaxaMoeda1.PromptInclude = False
        TaxaMoeda1.Text = Format(objDIInfo.dTaxaMoeda1, TaxaMoeda1.Format)
        TaxaMoeda1.PromptInclude = True
    End If
    
    For iIndice = 0 To Moeda2.ListCount - 1
        If objDIInfo.iMoeda2 = Codigo_Extrai(Moeda2.List(iIndice)) Then
            Moeda2.ListIndex = iIndice
            Exit For
        End If
    Next

    If objDIInfo.dTaxaMoeda2 <> 0 Then
        TaxaMoeda2.PromptInclude = False
        TaxaMoeda2.Text = Format(objDIInfo.dTaxaMoeda2, TaxaMoeda2.Format)
        TaxaMoeda2.PromptInclude = True
    End If
    
    For iIndice = 0 To DUF.ListCount - 1
        If objDIInfo.sUFDesembaraco = DUF.List(iIndice) Then
            DUF.ListIndex = iIndice
            Exit For
        End If
    Next

    DLocal.Text = objDIInfo.sLocalDesembaraco
    
    If objDIInfo.dtDataDesembaraco <> DATA_NULA Then
        DData.PromptInclude = False
        DData.Text = Format(objDIInfo.dtDataDesembaraco, "dd/mm/yy")
        DData.PromptInclude = True
    End If
    
    MoedaFrete.ListIndex = objDIInfo.iMoedaFrete - 1
    MoedaItens.ListIndex = objDIInfo.iMoedaItens - 1
    MoedaMercadoria.ListIndex = objDIInfo.iMoedaMercadoria - 1
    MoedaSeguro.ListIndex = objDIInfo.iMoedaSeguro - 1
    
    CodExportador.Text = objDIInfo.sCodExportador
    
    If objDIInfo.dPesoBrutoKG <> 0 Then
        DIPesoBruto.Text = Formata_Estoque(objDIInfo.dPesoBrutoKG)
    End If

    If objDIInfo.dPesoLiqKG <> 0 Then
        DIPesoLiquido.Text = Formata_Estoque(objDIInfo.dPesoLiqKG)
    End If
    
    Call Combo_Seleciona_ItemData(Intermedio, objDIInfo.iIntermedio)
    Call Combo_Seleciona_ItemData(ViaTransp, objDIInfo.iViaTransp)
    For iIndice = 0 To UFAdquir.ListCount - 1
        If objDIInfo.sUFAdquir = UFAdquir.List(iIndice) Then
            UFAdquir.ListIndex = iIndice
            Exit For
        End If
    Next
    If objDIInfo.sCNPJAdquir <> "" Then
        CNPJAdquir.Text = objDIInfo.sCNPJAdquir
        Call CNPJAdquir_Validate(bSGECancelDummy)
    Else
        CNPJAdquir.Text = ""
    End If

'    If objDIInfo.dValorMercadoriaMoeda <> 0 Then
'        ValorMercadoriaMoeda.PromptInclude = False
'        ValorMercadoriaMoeda.Text = Format(objDIInfo.dValorMercadoriaMoeda, ValorMercadoriaMoeda.Format)
'        ValorMercadoriaMoeda.PromptInclude = True
'    End If

    If objDIInfo.dValorFreteInternacMoeda <> 0 Then
        ValorFreteInternacMoeda.PromptInclude = False
        ValorFreteInternacMoeda.Text = Format(objDIInfo.dValorFreteInternacMoeda, ValorFreteInternacMoeda.Format)
        ValorFreteInternacMoeda.PromptInclude = True
    End If

    If objDIInfo.dValorSeguroInternacMoeda <> 0 Then
        ValorSeguroInternacMoeda.PromptInclude = False
        ValorSeguroInternacMoeda.Text = Format(objDIInfo.dValorSeguroInternacMoeda, ValorSeguroInternacMoeda.Format)
        ValorSeguroInternacMoeda.PromptInclude = True
    End If

'    If objDIInfo.dValorMercadoriaEmReal <> 0 Then
'        ValorMercadoriaEmReal.PromptInclude = False
'        ValorMercadoriaEmReal.Text = Format(objDIInfo.dValorMercadoriaEmReal, ValorMercadoriaEmReal.Format)
'        ValorMercadoriaEmReal.PromptInclude = True
'    End If
'
'    If objDIInfo.dValorFreteInternacEmReal <> 0 Then
'        ValorFreteInternacEmReal.PromptInclude = False
'        ValorFreteInternacEmReal.Text = Format(objDIInfo.dValorFreteInternacEmReal, ValorFreteInternacEmReal.Format)
'        ValorFreteInternacEmReal.PromptInclude = True
'    End If
'
'    If objDIInfo.dValorSeguroInternacEmReal <> 0 Then
'        ValorSeguroInternacEmReal.PromptInclude = False
'        ValorSeguroInternacEmReal.Text = Format(objDIInfo.dValorSeguroInternacEmReal, ValorSeguroInternacEmReal.Format)
'        ValorSeguroInternacEmReal.PromptInclude = True
'    End If
    
    lErro = Preenche_GridAdicao_Tela(objDIInfo)
    If lErro <> SUCESSO Then gError 196591
    
'    lErro = Preenche_GridDespesas_Tela(objDIInfo)
'    If lErro <> SUCESSO Then gError 196592
       
    iTaxaMoeda1Alterada = REGISTRO_ALTERADO
    Call TaxaMoeda1_Validate(bSGECancelDummy)

    iTaxaMoeda2Alterada = REGISTRO_ALTERADO
    Call TaxaMoeda2_Validate(bSGECancelDummy)

'    lErro = Preenche_GridItensPCDI_Tela(objDIInfo)
'    If lErro <> SUCESSO Then gError 210605

    Call Calcula_Valores
    
    ValorMercadoriaMoeda.PromptInclude = False
    ValorMercadoriaMoeda.Text = Format(StrParaDbl(LabelTotalFOBDIMoeda.Caption), ValorMercadoriaMoeda.Format)
    ValorMercadoriaMoeda.PromptInclude = True
    Call ValorMercadoriaMoeda_Validate(bSGECancelDummy)
    
    Traz_DIInfo_Tela_XML = SUCESSO

    Exit Function

Erro_Traz_DIInfo_Tela_XML:

    Traz_DIInfo_Tela_XML = gErr

    Select Case gErr

        Case 196590 To 196592, 210605

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196593)

    End Select
    
    Exit Function

End Function

Private Sub Exibe_CampoDet_Grid(ByVal objGridInt As AdmGrid, ByVal iColunaExibir As Integer, ByVal objControle As Object)

Dim iLinha As Integer

On Error GoTo Erro_Exibe_CampoDet_Grid

    iLinha = objGridInt.objGrid.Row
    
    If iLinha > 0 And iLinha <= objGridInt.iLinhasExistentes Then
        objControle.Caption = objGridInt.objGrid.TextMatrix(iLinha, iColunaExibir)
    Else
        objControle.Caption = ""
    End If

    Exit Sub

Erro_Exibe_CampoDet_Grid:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208641)

    End Select

    Exit Sub
    
End Sub
