VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl Produto 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4215
      Index           =   4
      Left            =   180
      TabIndex        =   108
      Top             =   1290
      Visible         =   0   'False
      Width           =   9240
      Begin VB.Frame Frame24 
         Caption         =   "Outras"
         Height          =   2325
         Left            =   90
         TabIndex        =   156
         Top             =   1665
         Width           =   9000
         Begin VB.TextBox ObsFisica 
            Height          =   1425
            Left            =   1290
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   60
            Top             =   750
            Width           =   7425
         End
         Begin VB.TextBox Cor 
            Height          =   300
            Left            =   1275
            MaxLength       =   20
            TabIndex        =   58
            Top             =   255
            Width           =   1995
         End
         Begin MSMask.MaskEdBox QtdePorEmb 
            Height          =   285
            Left            =   7050
            TabIndex        =   59
            Top             =   240
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Qtde por Embalagem:"
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
            TabIndex        =   216
            Top             =   285
            Width           =   1830
         End
         Begin VB.Label Label13 
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
            Left            =   135
            TabIndex        =   158
            Top             =   810
            Width           =   1095
         End
         Begin VB.Label Label12 
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
            Left            =   855
            TabIndex        =   157
            Top             =   315
            Width           =   360
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Medidas (em metros)"
         Height          =   720
         Left            =   90
         TabIndex        =   152
         Top             =   840
         Width           =   9015
         Begin MSMask.MaskEdBox Comprimento 
            Height          =   285
            Left            =   1275
            TabIndex        =   55
            Top             =   300
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Largura 
            Height          =   285
            Left            =   7050
            TabIndex        =   57
            Top             =   270
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Espessura 
            Height          =   285
            Left            =   4065
            TabIndex        =   56
            Top             =   285
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00######"
            PromptChar      =   " "
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Comprimento:"
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
            Left            =   75
            TabIndex        =   155
            Top             =   345
            Width           =   1155
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Largura:"
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
            TabIndex        =   154
            Top             =   330
            Width           =   720
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Espessura:"
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
            Left            =   3030
            TabIndex        =   153
            Top             =   330
            Width           =   945
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Peso"
         Height          =   675
         Left            =   90
         TabIndex        =   145
         Top             =   90
         Width           =   9015
         Begin MSMask.MaskEdBox PesoLiquido 
            Height          =   285
            Left            =   1290
            TabIndex        =   52
            Top             =   240
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PesoBruto 
            Height          =   285
            Left            =   4065
            TabIndex        =   53
            Top             =   255
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PesoEspecifico 
            Height          =   285
            Left            =   7020
            TabIndex        =   54
            Top             =   255
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00#"
            PromptChar      =   " "
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   525
            TabIndex        =   151
            Top             =   285
            Width           =   705
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   3465
            TabIndex        =   150
            Top             =   300
            Width           =   525
         End
         Begin VB.Label Label15 
            Caption         =   "Kg"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2610
            TabIndex        =   149
            Top             =   285
            Width           =   330
         End
         Begin VB.Label Label16 
            Caption         =   "Kg"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5370
            TabIndex        =   148
            Top             =   300
            Width           =   330
         End
         Begin VB.Label LabelPesoEspecifico 
            AutoSize        =   -1  'True
            Caption         =   "Específico:"
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
            Left            =   5910
            TabIndex        =   147
            Top             =   300
            Width           =   975
         End
         Begin VB.Label LabelPesoEspKg 
            Caption         =   "Kg/l"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   8325
            TabIndex        =   146
            Top             =   300
            Width           =   510
         End
      End
      Begin VB.TextBox Embalagem 
         Height          =   300
         Left            =   -1965
         MaxLength       =   20
         TabIndex        =   110
         Top             =   2295
         Width           =   540
      End
      Begin VB.Label DescricaoEmbalagem 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   -4500
         TabIndex        =   111
         Top             =   2295
         Width           =   3975
      End
      Begin VB.Label LabelEmbalagem 
         AutoSize        =   -1  'True
         Caption         =   "Embalagem Padrão:"
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
         Left            =   -2000
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   109
         Top             =   2355
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4125
      Index           =   6
      Left            =   195
      TabIndex        =   118
      Top             =   1320
      Visible         =   0   'False
      Width           =   9225
      Begin VB.Frame Frame28 
         Caption         =   "Demais Informações"
         Height          =   1425
         Left            =   0
         TabIndex        =   189
         Top             =   2535
         Width           =   9180
         Begin VB.ComboBox ProdutoEspecifico 
            Height          =   315
            ItemData        =   "produtonena.ctx":0000
            Left            =   2190
            List            =   "produtonena.ctx":0014
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   255
            Width           =   2925
         End
         Begin MSMask.MaskEdBox Genero 
            Height          =   315
            Left            =   2190
            TabIndex        =   73
            Top             =   660
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   2
            Format          =   "00"
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ISSQN 
            Height          =   315
            Left            =   2190
            TabIndex        =   74
            Top             =   1095
            Visible         =   0   'False
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Format          =   "0000"
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrecoMaxConsumidor 
            Height          =   315
            Left            =   7800
            TabIndex        =   72
            Top             =   270
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
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
            Format          =   "#,##0.0000"
            PromptChar      =   " "
         End
         Begin VB.Label LabelCodServNFe 
            AutoSize        =   -1  'True
            Caption         =   "Código Serviço NFe:"
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
            Left            =   300
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   212
            Top             =   1470
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.Label CodServNFe 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2190
            TabIndex        =   211
            Top             =   1410
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Preço máximo ao consumidor:"
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
            Left            =   5190
            TabIndex        =   197
            Top             =   300
            Width           =   2520
         End
         Begin VB.Label DescGenero 
            BorderStyle     =   1  'Fixed Single
            Height          =   690
            Left            =   2580
            TabIndex        =   196
            Top             =   660
            Width           =   6540
         End
         Begin VB.Label LabelGenero 
            AutoSize        =   -1  'True
            Caption         =   "Gênero:"
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
            Left            =   1395
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   195
            Top             =   720
            Width           =   675
         End
         Begin VB.Label LabelISSQN 
            AutoSize        =   -1  'True
            Caption         =   "ISSQN:"
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
            Left            =   1425
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   194
            Top             =   1140
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label DescISSQN 
            BorderStyle     =   1  'Fixed Single
            Height          =   630
            Left            =   2850
            TabIndex        =   193
            Top             =   1095
            Visible         =   0   'False
            Width           =   6285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Produto específico:"
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
            Left            =   390
            TabIndex        =   190
            Top             =   300
            Width           =   1680
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Compras"
         Height          =   2325
         Left            =   0
         TabIndex        =   179
         Top             =   105
         Width           =   9180
         Begin VB.Frame Frame10 
            Caption         =   "Cotações Anteriores"
            Height          =   1770
            Left            =   60
            TabIndex        =   185
            Top             =   315
            Width           =   3240
            Begin VB.Frame Frame11 
               Caption         =   "Limites % de quantidade para uso"
               Height          =   990
               Index           =   0
               Left            =   45
               TabIndex        =   186
               Top             =   570
               Width           =   3000
               Begin MSMask.MaskEdBox PercentMaisQuantCotacaoAnterior 
                  Height          =   315
                  Left            =   2085
                  TabIndex        =   64
                  Top             =   255
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
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
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox PercentMenosQuantCotacaoAnterior 
                  Height          =   315
                  Left            =   2085
                  TabIndex        =   65
                  Top             =   585
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
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
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin VB.Label Label25 
                  AutoSize        =   -1  'True
                  Caption         =   "Percentagem a mais:"
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
                  TabIndex        =   188
                  Top             =   315
                  Width           =   1785
               End
               Begin VB.Label Label26 
                  AutoSize        =   -1  'True
                  Caption         =   "Percentagem a menos:"
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
                  Left            =   45
                  TabIndex        =   187
                  Top             =   645
                  Width           =   1950
               End
            End
            Begin VB.CheckBox ConsideraQuantCotacaoAnterior 
               Caption         =   "Usa independente de quantidade"
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
               Left            =   45
               TabIndex        =   63
               Top             =   270
               Width           =   3165
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Recebimento"
            Height          =   1770
            Left            =   3330
            TabIndex        =   180
            Top             =   315
            Width           =   5790
            Begin VB.Frame Frame14 
               Caption         =   "Faixa de recebimento"
               Height          =   960
               Left            =   75
               TabIndex        =   182
               Top             =   600
               Width           =   2940
               Begin MSMask.MaskEdBox PercentMaisReceb 
                  Height          =   315
                  Left            =   2055
                  TabIndex        =   67
                  Top             =   240
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
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
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox PercentMenosReceb 
                  Height          =   315
                  Left            =   2055
                  TabIndex        =   68
                  Top             =   570
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
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
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  Caption         =   "Porcentagem a menos:"
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
                  Left            =   45
                  TabIndex        =   184
                  Top             =   630
                  Width           =   1950
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  Caption         =   "Porcentagem a mais:"
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
                  TabIndex        =   183
                  Top             =   300
                  Width           =   1785
               End
            End
            Begin VB.Frame Frame13 
               Caption         =   "Recebimento fora da faixa"
               Height          =   960
               Left            =   3030
               TabIndex        =   181
               Top             =   600
               Width           =   2715
               Begin VB.OptionButton RecebForaFaixa 
                  Caption         =   "Não aceita recebimento"
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
                  Index           =   0
                  Left            =   30
                  TabIndex        =   69
                  Top             =   270
                  Value           =   -1  'True
                  Width           =   2415
               End
               Begin VB.OptionButton RecebForaFaixa 
                  Caption         =   "Avisa e aceita recebimento"
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
                  Index           =   1
                  Left            =   30
                  TabIndex        =   70
                  Top             =   570
                  Width           =   2655
               End
            End
            Begin VB.CheckBox NaoTemFaixaReceb 
               Caption         =   "Aceita qualquer quantidade sem aviso"
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
               Left            =   75
               TabIndex        =   66
               Top             =   285
               Width           =   3585
            End
         End
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   4110
      Index           =   7
      Left            =   120
      TabIndex        =   124
      Top             =   1335
      Visible         =   0   'False
      Width           =   9285
      Begin VB.ComboBox OrigemMercadoria 
         Height          =   315
         ItemData        =   "produtonena.ctx":0059
         Left            =   4500
         List            =   "produtonena.ctx":0075
         Style           =   2  'Dropdown List
         TabIndex        =   214
         Top             =   435
         Width           =   4515
      End
      Begin VB.Frame Frame29 
         Caption         =   "NCM"
         Height          =   750
         Left            =   4725
         TabIndex        =   191
         Top             =   2310
         Width           =   3990
         Begin MSMask.MaskEdBox ClasFiscIPI 
            Height          =   300
            Left            =   1935
            TabIndex        =   83
            Top             =   255
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "0000\.00\.00"
            Mask            =   "##########"
            PromptChar      =   " "
         End
         Begin VB.Label LabelClassificacaoFiscal 
            AutoSize        =   -1  'True
            Caption         =   "Classificação Fiscal:"
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
            Left            =   135
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   192
            Top             =   315
            Width           =   1755
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "INSS"
         Height          =   750
         Left            =   240
         TabIndex        =   135
         Top             =   3075
         Width           =   4035
         Begin MSMask.MaskEdBox INSSPercBase 
            Height          =   285
            Left            =   1980
            TabIndex        =   80
            Top             =   255
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "% da Base de Cálculo:"
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
            Left            =   15
            TabIndex        =   136
            Top             =   285
            Width           =   1920
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Loja"
         Height          =   1500
         Left            =   225
         TabIndex        =   132
         Top             =   0
         Width           =   4050
         Begin VB.CheckBox UsaBalanca 
            Caption         =   "Usa Balança"
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
            Left            =   1965
            TabIndex        =   77
            Top             =   1155
            Width           =   1860
         End
         Begin VB.ComboBox comboAliquota 
            Height          =   315
            ItemData        =   "produtonena.ctx":0371
            Left            =   1980
            List            =   "produtonena.ctx":0373
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   720
            Width           =   1500
         End
         Begin VB.ComboBox SituacaoTributaria 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Top             =   345
            Width           =   1950
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Situação Tributária:"
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
            TabIndex        =   134
            Top             =   375
            Width           =   1695
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Alíquota:"
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
            Left            =   1140
            TabIndex        =   133
            Top             =   780
            Width           =   795
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "IPI"
         Height          =   2130
         Left            =   4470
         TabIndex        =   125
         Top             =   1695
         Width           =   4530
         Begin VB.Frame Frame30 
            Caption         =   "Exceção a TIPI"
            Height          =   645
            Left            =   255
            TabIndex        =   198
            Top             =   1380
            Width           =   3990
            Begin MSMask.MaskEdBox ExTIPI 
               Height          =   300
               Left            =   1935
               TabIndex        =   84
               Top             =   255
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   3
               Mask            =   "###"
               PromptChar      =   " "
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Código:"
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
               Left            =   1215
               TabIndex        =   199
               Top             =   315
               Width           =   660
            End
         End
         Begin VB.CheckBox IncideIPI 
            Caption         =   "Incide"
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
            Left            =   255
            TabIndex        =   81
            Top             =   270
            Value           =   1  'Checked
            Width           =   915
         End
         Begin MSMask.MaskEdBox AliquotaIPI 
            Height          =   285
            Left            =   2190
            TabIndex        =   82
            Top             =   255
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodigoIPI 
            Height          =   300
            Left            =   3870
            TabIndex        =   117
            Top             =   240
            Visible         =   0   'False
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
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
            Left            =   3150
            TabIndex        =   127
            Top             =   300
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Alíquota:"
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
            Left            =   1320
            TabIndex        =   126
            Top             =   270
            Width           =   795
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Contabilidade"
         Height          =   1185
         Left            =   225
         TabIndex        =   128
         Top             =   1695
         Width           =   4050
         Begin MSMask.MaskEdBox ContaContabil 
            Height          =   315
            Left            =   1980
            TabIndex        =   78
            Top             =   225
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ContaProducao 
            Height          =   315
            Left            =   1980
            TabIndex        =   79
            Top             =   660
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LabelContaProducao 
            AutoSize        =   -1  'True
            Caption         =   "Conta de Produção:"
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
            Left            =   210
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   130
            ToolTipText     =   "Conta Contábil de Produção"
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label ContaContabilLabel 
            AutoSize        =   -1  'True
            Caption         =   "Conta de Aplicação:"
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
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   129
            ToolTipText     =   "Conta Contábil de Aplicação"
            Top             =   285
            Width           =   1755
         End
      End
      Begin VB.Label LabelOrigemMerc 
         Caption         =   "Orgem da Mercadoria:"
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
         Left            =   4485
         TabIndex        =   215
         Top             =   120
         Width           =   2520
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4215
      Index           =   1
      Left            =   165
      TabIndex        =   200
      Top             =   1200
      Width           =   9195
      Begin VB.ComboBox Fabricante 
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3120
         Width           =   3345
      End
      Begin VB.TextBox Descricao 
         Height          =   312
         Left            =   1470
         MaxLength       =   20
         TabIndex        =   3
         Top             =   765
         Width           =   4035
      End
      Begin VB.CheckBox Ativo 
         Caption         =   "Ativo"
         Height          =   375
         Left            =   4320
         TabIndex        =   2
         Top             =   315
         Value           =   1  'Checked
         Width           =   705
      End
      Begin VB.TextBox Modelo 
         Height          =   312
         Left            =   1470
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1935
         Width           =   2355
      End
      Begin VB.TextBox NomeReduzido 
         Height          =   312
         Left            =   1470
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1140
         Width           =   2355
      End
      Begin VB.ComboBox NaturezaProduto 
         Height          =   315
         ItemData        =   "produtonena.ctx":0375
         Left            =   1470
         List            =   "produtonena.ctx":03A4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2715
         Width           =   3360
      End
      Begin VB.CommandButton BotaoProcurar 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8355
         TabIndex        =   13
         Top             =   375
         Width           =   300
      End
      Begin VB.CommandButton BotaoVisualizar 
         Caption         =   "Visualizar"
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
         Left            =   6495
         TabIndex        =   14
         Top             =   750
         Width           =   1275
      End
      Begin VB.Frame Frame1 
         Caption         =   "Nível"
         Height          =   645
         Index           =   1
         Left            =   555
         TabIndex        =   201
         Top             =   3450
         Width           =   4215
         Begin VB.OptionButton NivelFinal 
            Caption         =   "Analítico"
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
            Left            =   2430
            TabIndex        =   11
            Top             =   270
            Width           =   1545
         End
         Begin VB.OptionButton NivelGerencial 
            Caption         =   "Gerencial"
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
            Left            =   735
            TabIndex        =   10
            Top             =   270
            Value           =   -1  'True
            Width           =   1245
         End
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   3840
         Picture         =   "produtonena.ctx":0497
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Numeração Automática"
         Top             =   375
         Width           =   300
      End
      Begin VB.TextBox Referencia 
         Height          =   312
         Left            =   1470
         TabIndex        =   5
         Top             =   1545
         Width           =   2355
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1470
         TabIndex        =   0
         Top             =   375
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TipoProduto 
         Height          =   315
         Left            =   1470
         TabIndex        =   7
         Top             =   2325
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeFigura 
         Height          =   315
         Left            =   5790
         TabIndex        =   12
         Top             =   375
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSComDlg.CommonDialog CommonDialog 
         Left            =   4530
         Top             =   1425
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Escolhendo Figura para o Produto"
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Fabricante:"
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
         Left            =   405
         TabIndex        =   213
         Top             =   3180
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Modelo:"
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
         Left            =   720
         TabIndex        =   210
         Top             =   1980
         Width           =   690
      End
      Begin VB.Label LabelNomeReduzido 
         AutoSize        =   -1  'True
         Caption         =   "Nome Reduzido:"
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
         Left            =   15
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   209
         Top             =   1200
         Width           =   1410
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   208
         Top             =   795
         Width           =   930
      End
      Begin VB.Label LabelCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
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
         Left            =   750
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   207
         Top             =   390
         Width           =   660
      End
      Begin VB.Label LblTipoProduto 
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
         Left            =   975
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   206
         Top             =   2385
         Width           =   450
      End
      Begin VB.Label DescTipoProduto 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2130
         TabIndex        =   205
         Top             =   2325
         Width           =   3015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Natureza:"
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
         Left            =   570
         TabIndex        =   204
         Top             =   2745
         Width           =   840
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Referência:"
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
         TabIndex        =   203
         Top             =   1590
         Width           =   1005
      End
      Begin VB.Image Figura 
         BorderStyle     =   1  'Fixed Single
         Height          =   2970
         Left            =   5610
         Stretch         =   -1  'True
         Top             =   1125
         Width           =   3030
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Figura:"
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
         Index           =   2
         Left            =   5130
         TabIndex        =   202
         Top             =   420
         Width           =   600
      End
   End
   Begin VB.CommandButton BotaoSRV 
      Caption         =   "Serviço"
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
      Left            =   8160
      TabIndex        =   21
      Top             =   5550
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4245
      Index           =   5
      Left            =   75
      TabIndex        =   119
      Top             =   1260
      Visible         =   0   'False
      Width           =   9300
      Begin VB.TextBox DescricaoTabela 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   2310
         TabIndex        =   113
         Text            =   "DescricaoTabela"
         Top             =   405
         Width           =   3030
      End
      Begin VB.TextBox Tabela 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   1455
         TabIndex        =   112
         Text            =   "Tabela"
         Top             =   405
         Width           =   1125
      End
      Begin VB.CommandButton BotaoTabelaPreco 
         Caption         =   "Tabela de Preços"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7095
         TabIndex        =   62
         Top             =   3780
         Width           =   2025
      End
      Begin MSMask.MaskEdBox ValorFilial 
         Height          =   225
         Left            =   5880
         TabIndex        =   115
         Top             =   390
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
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorEmpresa 
         Height          =   225
         Left            =   4575
         TabIndex        =   114
         Top             =   390
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
      Begin MSMask.MaskEdBox DataPreco 
         Height          =   225
         Left            =   7140
         TabIndex        =   116
         Tag             =   "1"
         Top             =   390
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
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
      Begin MSFlexGridLib.MSFlexGrid GridTabelaPreco 
         Height          =   3360
         Left            =   225
         TabIndex        =   61
         Top             =   345
         Width           =   8940
         _ExtentX        =   15769
         _ExtentY        =   5927
         _Version        =   393216
         Rows            =   11
         Cols            =   5
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   0
      End
      Begin VB.Label DescrUM 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2685
         TabIndex        =   122
         Top             =   3765
         Width           =   1665
      End
      Begin VB.Label Label4 
         Caption         =   "Unidade Medida de Venda:"
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
         Left            =   225
         TabIndex        =   121
         Top             =   3840
         Width           =   2355
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Tabelas de Preço de Venda"
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
         Left            =   255
         TabIndex        =   120
         Top             =   75
         Width           =   2385
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame10"
      Height          =   4110
      Index           =   2
      Left            =   150
      TabIndex        =   86
      Top             =   1290
      Visible         =   0   'False
      Width           =   9240
      Begin VB.Frame Frame16 
         Caption         =   "Unidade de Medida"
         Height          =   2475
         Left            =   15
         TabIndex        =   159
         Top             =   -30
         Width           =   4410
         Begin VB.Frame Frame17 
            Caption         =   "Unidade Padrão"
            Height          =   1755
            Left            =   75
            TabIndex        =   160
            Top             =   615
            Width           =   4245
            Begin VB.ComboBox SiglaUMTrib 
               Height          =   315
               Left            =   990
               Style           =   2  'Dropdown List
               TabIndex        =   30
               Top             =   1365
               Width           =   915
            End
            Begin VB.ComboBox SiglaUMEstoque 
               Height          =   315
               Left            =   990
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Top             =   240
               Width           =   915
            End
            Begin VB.ComboBox SiglaUMCompra 
               Height          =   315
               Left            =   990
               Style           =   2  'Dropdown List
               TabIndex        =   28
               Top             =   615
               Width           =   915
            End
            Begin VB.ComboBox SiglaUMVenda 
               Height          =   315
               Left            =   990
               Style           =   2  'Dropdown List
               TabIndex        =   29
               Top             =   990
               Width           =   915
            End
            Begin VB.Label NomeUMTrib 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1905
               TabIndex        =   178
               Top             =   1365
               Width           =   2280
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               Caption         =   "Tributável:"
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
               Left            =   30
               TabIndex        =   177
               Top             =   1410
               Width           =   915
            End
            Begin VB.Label LblUMEstoque 
               AutoSize        =   -1  'True
               Caption         =   "Estoque:"
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
               Left            =   180
               TabIndex        =   166
               Top             =   270
               Width           =   765
            End
            Begin VB.Label NomeUMEstoque 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1905
               TabIndex        =   165
               Top             =   240
               Width           =   2280
            End
            Begin VB.Label LblUMCompras 
               AutoSize        =   -1  'True
               Caption         =   "Compras:"
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
               Left            =   150
               TabIndex        =   164
               Top             =   660
               Width           =   795
            End
            Begin VB.Label NomeUMCompra 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1905
               TabIndex        =   163
               Top             =   615
               Width           =   2280
            End
            Begin VB.Label LblUMVendas 
               AutoSize        =   -1  'True
               Caption         =   "Vendas:"
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
               Left            =   225
               TabIndex        =   162
               Top             =   1035
               Width           =   705
            End
            Begin VB.Label NomeUMVenda 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1905
               TabIndex        =   161
               Top             =   990
               Width           =   2280
            End
         End
         Begin MSMask.MaskEdBox ClasseUM 
            Height          =   315
            Left            =   1065
            TabIndex        =   26
            Top             =   270
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "9999"
            PromptChar      =   " "
         End
         Begin VB.Label DescricaoClasseUM 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1560
            TabIndex        =   168
            Top             =   270
            Width           =   2670
         End
         Begin VB.Label LblClasseUM 
            AutoSize        =   -1  'True
            Caption         =   "Classe:"
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
            Left            =   375
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   167
            Top             =   300
            Width           =   630
         End
      End
      Begin VB.Frame FrameGrade 
         Caption         =   "Grade"
         Height          =   690
         Left            =   4575
         TabIndex        =   138
         Top             =   -30
         Width           =   4620
         Begin VB.ComboBox Grades 
            Height          =   315
            Left            =   900
            TabIndex        =   34
            Top             =   270
            Width           =   2010
         End
         Begin VB.CommandButton BotaoCriarGrade 
            Caption         =   "Criar Grade "
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
            Left            =   2925
            TabIndex        =   35
            Top             =   255
            Width           =   1260
         End
         Begin VB.Label LabelGrade 
            AutoSize        =   -1  'True
            Caption         =   "Grade:"
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
            Left            =   210
            TabIndex        =   139
            Top             =   315
            Width           =   585
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Categorias"
         Height          =   3345
         Left            =   4575
         TabIndex        =   88
         Top             =   780
         Width           =   4620
         Begin VB.ComboBox ComboCategoriaProduto 
            Height          =   315
            Left            =   570
            TabIndex        =   89
            Top             =   540
            Width           =   1590
         End
         Begin VB.ComboBox ComboCategoriaProdutoItem 
            Height          =   315
            Left            =   2025
            TabIndex        =   90
            Top             =   540
            Width           =   2190
         End
         Begin MSFlexGridLib.MSFlexGrid GridCategoria 
            Height          =   3000
            Left            =   45
            TabIndex        =   36
            Top             =   300
            Width           =   4530
            _ExtentX        =   7990
            _ExtentY        =   5292
            _Version        =   393216
            Rows            =   6
            Cols            =   3
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Características"
         Height          =   1680
         Index           =   1
         Left            =   15
         TabIndex        =   87
         Top             =   2445
         Width           =   4410
         Begin VB.ListBox ListaCaracteristicas 
            Height          =   1185
            ItemData        =   "produtonena.ctx":0581
            Left            =   255
            List            =   "produtonena.ctx":0594
            Style           =   1  'Checkbox
            TabIndex        =   31
            Top             =   195
            Width           =   3885
         End
         Begin VB.OptionButton Produzido 
            Caption         =   "Produzido"
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
            Left            =   750
            TabIndex        =   32
            Top             =   1395
            Width           =   1395
         End
         Begin VB.OptionButton Comprado 
            Caption         =   "Comprado"
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
            Left            =   2400
            TabIndex        =   33
            Top             =   1425
            Value           =   -1  'True
            Width           =   1245
         End
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      ForeColor       =   &H00000080&
      Height          =   4230
      Index           =   3
      Left            =   75
      TabIndex        =   91
      Top             =   1275
      Visible         =   0   'False
      Width           =   9360
      Begin VB.Frame Frame26 
         Caption         =   "Rastreamento"
         Height          =   570
         Left            =   135
         TabIndex        =   171
         Top             =   945
         Width           =   4125
         Begin VB.ComboBox Rastro 
            Height          =   315
            ItemData        =   "produtonena.ctx":062A
            Left            =   1470
            List            =   "produtonena.ctx":063A
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   180
            Width           =   2430
         End
         Begin VB.Label LabelRastro 
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
            Left            =   915
            TabIndex        =   172
            Top             =   225
            Width           =   450
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Geração de Número de Série"
         Height          =   1095
         Left            =   135
         TabIndex        =   140
         Top             =   2940
         Width           =   4125
         Begin MSMask.MaskEdBox SerieProx 
            Height          =   315
            Left            =   1500
            TabIndex        =   44
            Top             =   285
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
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
         Begin MSMask.MaskEdBox SerieNum 
            Height          =   315
            Left            =   1500
            TabIndex        =   45
            Top             =   645
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##"
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin VB.Label SeriePartNum 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1920
            TabIndex        =   143
            Top             =   645
            Width           =   2025
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Parte Numérica:"
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
            Left            =   45
            TabIndex        =   142
            Top             =   705
            Width           =   1380
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Próx Núm Série:"
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
            Left            =   45
            TabIndex        =   141
            Top             =   345
            Width           =   1365
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Código de Barras"
         Height          =   1275
         Left            =   135
         TabIndex        =   97
         Top             =   1590
         Width           =   4110
         Begin VB.Frame Frame25 
            Caption         =   "Número de Etiquetas Impressas"
            Height          =   585
            Left            =   420
            TabIndex        =   169
            Top             =   570
            Width           =   2715
            Begin MSMask.MaskEdBox EtiquetasCodBarras 
               Height          =   315
               Left            =   1050
               TabIndex        =   43
               Top             =   210
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "Etiquetas:"
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
               Left            =   75
               TabIndex        =   170
               Top             =   255
               Width           =   870
            End
         End
         Begin VB.CommandButton BotaoProxCodBarras 
            Height          =   285
            Left            =   3195
            Picture         =   "produtonena.ctx":0663
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Numeração Automática"
            Top             =   255
            Width           =   300
         End
         Begin VB.ComboBox CodigoBarras 
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton BotaoProdutoCodBarras 
            Caption         =   "..."
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
            Left            =   3570
            TabIndex        =   42
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
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
            TabIndex        =   98
            Top             =   285
            Width           =   660
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Produtos Substitutos"
         Height          =   915
         Left            =   120
         TabIndex        =   92
         Top             =   -30
         Width           =   9135
         Begin MSMask.MaskEdBox Substituto1 
            Height          =   315
            Left            =   1485
            TabIndex        =   37
            Top             =   210
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Substituto2 
            Height          =   315
            Left            =   1485
            TabIndex        =   38
            Top             =   540
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label DescSubst2 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3720
            TabIndex        =   96
            Top             =   525
            Width           =   5250
         End
         Begin VB.Label DescSubst1 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3720
            TabIndex        =   94
            Top             =   210
            Width           =   5250
         End
         Begin VB.Label LblSubst2 
            AutoSize        =   -1  'True
            Caption         =   "Produto 2:"
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
            Left            =   480
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   95
            Top             =   600
            Width           =   900
         End
         Begin VB.Label LblSubst1 
            AutoSize        =   -1  'True
            Caption         =   "Produto 1:"
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
            Left            =   480
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   93
            Top             =   255
            Width           =   900
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Estatística"
         Height          =   630
         Left            =   4365
         TabIndex        =   99
         Top             =   3405
         Width           =   4875
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade em Pedido:"
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
            TabIndex        =   100
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label QuantPedido 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2160
            TabIndex        =   101
            Top             =   225
            Width           =   1500
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Produção / Custo"
         Height          =   2400
         Left            =   4365
         TabIndex        =   102
         Top             =   945
         Width           =   4890
         Begin VB.ComboBox ApropriacaoComp 
            Height          =   315
            ItemData        =   "produtonena.ctx":074D
            Left            =   2130
            List            =   "produtonena.ctx":0754
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   195
            Width           =   2610
         End
         Begin VB.ComboBox ApropriacaoProd 
            Height          =   315
            ItemData        =   "produtonena.ctx":0765
            Left            =   2130
            List            =   "produtonena.ctx":076F
            Style           =   2  'Dropdown List
            TabIndex        =   123
            Top             =   195
            Visible         =   0   'False
            Width           =   2610
         End
         Begin MSMask.MaskEdBox CustoReposicao 
            Height          =   315
            Left            =   2130
            TabIndex        =   49
            Top             =   1257
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
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
            Format          =   "#,##0.0000"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrazoValidade 
            Height          =   315
            Left            =   2130
            TabIndex        =   47
            Top             =   549
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Residuo 
            Height          =   315
            Left            =   2130
            TabIndex        =   48
            ToolTipText     =   "Percentagem máxima para Requisição ou Pedido de Compras poder ser baixado por resíduo."
            Top             =   903
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TempoProducao 
            Height          =   315
            Left            =   2130
            TabIndex        =   50
            Top             =   1611
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox HorasMaquina 
            Height          =   315
            Left            =   2130
            TabIndex        =   51
            Top             =   1965
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "dias"
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
            Left            =   2745
            TabIndex        =   176
            Top             =   1650
            Width           =   360
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "dias"
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
            TabIndex        =   175
            Top             =   615
            Width           =   360
         End
         Begin VB.Label LabelHorasMaq 
            AutoSize        =   -1  'True
            Caption         =   "Horas de Máquina:"
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
            Left            =   405
            TabIndex        =   174
            Top             =   2010
            Width           =   1620
         End
         Begin VB.Label LabelMinutos 
            Caption         =   "minutos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2805
            TabIndex        =   173
            Top             =   2010
            Width           =   810
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Tempo de Produção:"
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
            Left            =   255
            TabIndex        =   107
            Top             =   1650
            Width           =   1770
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Prazo de Validade:"
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
            Left            =   405
            TabIndex        =   104
            Top             =   600
            Width           =   1620
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Apropriação:"
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
            Left            =   930
            TabIndex        =   103
            Top             =   225
            Width           =   1095
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Resíduo (%):"
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
            TabIndex        =   105
            Top             =   915
            Width           =   1110
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Custo de Reposição:"
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
            TabIndex        =   106
            Top             =   1275
            Width           =   1785
         End
      End
   End
   Begin VB.CommandButton BotaoCorTonTP 
      Caption         =   "Cadastrar Cor, Tonalidade e Tipo de Pintura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2730
      TabIndex        =   144
      Top             =   135
      Visible         =   0   'False
      Width           =   3930
   End
   Begin VB.CommandButton BotaoTeste 
      Caption         =   "Qualidade"
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
      Left            =   6810
      TabIndex        =   20
      Top             =   5550
      Width           =   1275
   End
   Begin VB.CommandButton BotaoEmbalagem 
      Caption         =   "Embalagens"
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
      Left            =   5295
      TabIndex        =   19
      Top             =   5550
      Width           =   1440
   End
   Begin VB.CommandButton BotaoFornecedores 
      Caption         =   "Fornecedores"
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
      Left            =   3765
      TabIndex        =   18
      Top             =   5565
      Width           =   1455
   End
   Begin VB.CommandButton BotaoEstoque 
      Caption         =   "Estoque"
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
      Left            =   2700
      TabIndex        =   17
      Top             =   5565
      Width           =   960
   End
   Begin VB.CommandButton BotaoCustos 
      Caption         =   "Custos"
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
      Left            =   1785
      TabIndex        =   16
      Top             =   5565
      Width           =   840
   End
   Begin VB.CommandButton BotaoControleEstoque 
      Caption         =   "Controle Estoque"
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
      TabIndex        =   15
      Top             =   5565
      Width           =   1665
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7320
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   131
      TabStop         =   0   'False
      Top             =   45
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "produtonena.ctx":079B
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "produtonena.ctx":0919
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "produtonena.ctx":0E4B
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "produtonena.ctx":0FD5
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4935
      Left            =   45
      TabIndex        =   85
      Top             =   600
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   8705
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "UM / Categoria / Grade"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Complemento"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Características Físicas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Preços"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Outros"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tributação / Contabilização"
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Alíquota ICMS:"
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
      Left            =   1035
      TabIndex        =   137
      Top             =   2460
      Width           =   1305
   End
End
Attribute VB_Name = "Produto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTProduto
Attribute objCT.VB_VarHelpID = -1

Private Sub OrigemMercadoria_Click()
     Call objCT.OrigemMercadoria_Click
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTProduto
    Set objCT.objUserControl = Me
    
    '#########################################
    'Nena
    Set objCT.gobjInfoUsu = New CTProdutoVGNena
    Set objCT.gobjInfoUsu.gobjTelaUsu = New CTProdutoNena
    '#########################################

End Sub

Private Sub BotaoCorTonTP_Click()
     Call objCT.BotaoCorTonTP_Click
End Sub

Private Sub BotaoCriarGrade_Click()
     Call objCT.BotaoCriarGrade_Click
End Sub

Private Sub BotaoProdutoCodBarras_Click()
     Call objCT.BotaoProdutoCodBarras_Click
End Sub

Private Sub Codigo_GotFocus()
     Call objCT.Codigo_GotFocus
End Sub

Private Sub comboAliquota_Change()
     Call objCT.comboAliquota_Change
End Sub

Private Sub comboAliquota_Click()
     Call objCT.comboAliquota_Click
End Sub

Private Sub BotaoEmbalagem_Click()
     Call objCT.BotaoEmbalagem_Click
End Sub

Private Sub BotaoProcurar_Click()
     Call objCT.BotaoProcurar_Click
End Sub

Private Sub BotaoVisualizar_Click()
     Call objCT.BotaoVisualizar_Click
End Sub

Private Sub NomeFigura_Change()
     Call objCT.NomeFigura_Change
End Sub

Private Sub LabelCodigo_Click()
     Call objCT.LabelCodigo_Click
End Sub

Private Sub LabelNomeReduzido_Click()
     Call objCT.LabelNomeReduzido_Click
End Sub

Private Sub ConsideraQuantCotacaoAnterior_Click()
     Call objCT.ConsideraQuantCotacaoAnterior_Click
End Sub

Private Sub CustoReposicao_Change()
     Call objCT.CustoReposicao_Change
End Sub

Private Sub CustoReposicao_LostFocus()
     Call objCT.CustoReposicao_LostFocus
End Sub

Private Sub BotaoControleEstoque_Click()
     Call objCT.BotaoControleEstoque_Click
End Sub

Private Sub Comprado_Click()
     Call objCT.Comprado_Click
End Sub

Private Sub ContaContabil_Change()
     Call objCT.ContaContabil_Change
End Sub

Private Sub ContaContabil_LostFocus()
     Call objCT.ContaContabil_LostFocus
End Sub

Private Sub ContaContabilLabel_Click()
     Call objCT.ContaContabilLabel_Click
End Sub

Private Sub ContaProducao_Change()
     Call objCT.ContaProducao_Change
End Sub

Private Sub ContaProducao_LostFocus()
     Call objCT.ContaProducao_LostFocus
End Sub

Private Sub AliquotaIPI_Change()
     Call objCT.AliquotaIPI_Change
End Sub

Private Sub AliquotaIPI_LostFocus()
     Call objCT.AliquotaIPI_LostFocus
End Sub

Private Sub Ativo_Click()
     Call objCT.Ativo_Click
End Sub

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub ClasFiscIPI_Change()
     Call objCT.ClasFiscIPI_Change
End Sub

Private Sub ClasFiscIPI_GotFocus()
     Call objCT.ClasFiscIPI_GotFocus
End Sub

Private Sub ClasseUM_Change()
     Call objCT.ClasseUM_Change
End Sub

Private Sub ClasseUM_GotFocus()
     Call objCT.ClasseUM_GotFocus
End Sub

Private Sub Codigo_Change()
     Call objCT.Codigo_Change
End Sub

Private Sub CodigoBarras_Change()
     Call objCT.CodigoBarras_Change
End Sub

Private Sub CodigoIPI_Change()
     Call objCT.CodigoIPI_Change
End Sub

Private Sub Comprimento_Change()
     Call objCT.Comprimento_Change
End Sub

Private Sub Comprimento_LostFocus()
     Call objCT.Comprimento_LostFocus
End Sub

Private Sub Descricao_Change()
     Call objCT.Descricao_Change
End Sub

Private Sub Espessura_Change()
     Call objCT.Espessura_Change
End Sub

Private Sub Espessura_LostFocus()
     Call objCT.Espessura_LostFocus
End Sub

Private Sub EtiquetasCodBarras_Change()
     Call objCT.EtiquetasCodBarras_Change
End Sub

Private Sub EtiquetasCodBarras_GotFocus()
     Call objCT.EtiquetasCodBarras_GotFocus
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub HorasMaquina_Change()
     Call objCT.HorasMaquina_Change
End Sub

Private Sub HorasMaquina_GotFocus()
     Call objCT.HorasMaquina_GotFocus
End Sub

Private Sub IncideIPI_Click()
     Call objCT.IncideIPI_Click
End Sub

Private Sub LabelContaProducao_Click()
     Call objCT.LabelContaProducao_Click
End Sub

Private Sub Largura_Change()
     Call objCT.Largura_Change
End Sub

Private Sub Largura_LostFocus()
     Call objCT.Largura_LostFocus
End Sub

Private Sub LblSubst1_Click()
     Call objCT.LblSubst1_Click
End Sub

Private Sub LblSubst2_Click()
     Call objCT.LblSubst2_Click
End Sub

Private Sub ListaCaracteristicas_Click()
     Call objCT.ListaCaracteristicas_Click
End Sub

Private Sub Modelo_Change()
     Call objCT.Modelo_Change
End Sub

Private Sub NaoTemFaixaReceb_Click()
     Call objCT.NaoTemFaixaReceb_Click
End Sub

Private Sub NaturezaProduto_Change()
     Call objCT.NaturezaProduto_Change
End Sub

Private Sub NivelFinal_LostFocus()
     Call objCT.NivelFinal_LostFocus
End Sub

Private Sub NivelGerencial_Click()
     Call objCT.NivelGerencial_Click
End Sub

Private Sub NomeReduzido_Change()
     Call objCT.NomeReduzido_Change
End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)
     Call objCT.NomeReduzido_Validate(Cancel)
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Public Function Trata_Parametros(Optional objProduto As ClassProduto) As Long
     Trata_Parametros = objCT.Trata_Parametros(objProduto)
End Function

Private Sub BotaoCustos_Click()
     Call objCT.BotaoCustos_Click
End Sub

Private Sub BotaoEstoque_Click()
     Call objCT.BotaoEstoque_Click
End Sub

Private Sub BotaoFornecedores_Click()
     Call objCT.BotaoFornecedores_Click
End Sub

Private Sub BotaoTabelaPreco_Click()
     Call objCT.BotaoTabelaPreco_Click
End Sub

Private Sub INSSPercBase_Change()
     Call objCT.INSSPercBase_Change
End Sub

Private Sub INSSPercBase_Validate(Cancel As Boolean)
     Call objCT.INSSPercBase_Validate(Cancel)
End Sub

Private Sub PesoBruto_Change()
     Call objCT.PesoBruto_Change
End Sub

Private Sub PesoBruto_LostFocus()
     Call objCT.PesoBruto_LostFocus
End Sub

Private Sub PesoEspecifico_Change()
     Call objCT.PesoEspecifico_Change
End Sub

Private Sub PesoEspecifico_Validate(Cancel As Boolean)
     Call objCT.PesoEspecifico_Validate(Cancel)
End Sub

Private Sub PesoLiquido_Change()
     Call objCT.PesoLiquido_Change
End Sub

Private Sub PesoLiquido_LostFocus()
     Call objCT.PesoLiquido_LostFocus
End Sub

Private Sub PrazoValidade_Change()
     Call objCT.PrazoValidade_Change
End Sub

Private Sub PrazoValidade_GotFocus()
     Call objCT.PrazoValidade_GotFocus
End Sub

Private Sub Produzido_Click()
     Call objCT.Produzido_Click
End Sub

Private Sub Referencia_Change()
     Call objCT.Referencia_Change
End Sub

Private Sub Residuo_Change()
     Call objCT.Residuo_Change
End Sub

Private Sub Residuo_LostFocus()
     Call objCT.Residuo_LostFocus
End Sub

Private Sub SiglaUMCompra_Click()
     Call objCT.SiglaUMCompra_Click
End Sub

Private Sub SiglaUMVenda_Click()
     Call objCT.SiglaUMVenda_Click
End Sub

Private Sub SiglaUMEstoque_Click()
     Call objCT.SiglaUMEstoque_Click
End Sub

Private Sub SituacaoTributaria_Click()
     Call objCT.SituacaoTributaria_Click
End Sub

Private Sub Substituto1_Validate(Cancel As Boolean)
     Call objCT.Substituto1_Validate(Cancel)
End Sub

Private Sub Substituto2_Validate(Cancel As Boolean)
     Call objCT.Substituto2_Validate(Cancel)
End Sub

Private Sub TempoProducao_Change()
     Call objCT.TempoProducao_Change
End Sub

Private Sub TempoProducao_GotFocus()
     Call objCT.TempoProducao_GotFocus
End Sub

Private Sub TipoProduto_Change()
     Call objCT.TipoProduto_Change
End Sub

Private Sub TipoProduto_GotFocus()
     Call objCT.TipoProduto_GotFocus
End Sub

Private Sub TipoProduto_Validate(Cancel As Boolean)
     Call objCT.TipoProduto_Validate(Cancel)
End Sub

Private Sub ClasseUM_Validate(Cancel As Boolean)
     Call objCT.ClasseUM_Validate(Cancel)
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)
     Call objCT.Codigo_Validate(Cancel)
End Sub

Private Sub LblTipoProduto_Click()
     Call objCT.LblTipoProduto_Click
End Sub

Private Sub LblClasseUM_Click()
     Call objCT.LblClasseUM_Click
End Sub

Private Sub NivelFinal_Click()
     Call objCT.NivelFinal_Click
End Sub

Private Sub Substituto1_Change()
     Call objCT.Substituto1_Change
End Sub

Private Sub Substituto2_Change()
     Call objCT.Substituto2_Change
End Sub

Private Sub GridCategoria_Click()
     Call objCT.GridCategoria_Click
End Sub

Private Sub GridCategoria_GotFocus()
     Call objCT.GridCategoria_GotFocus
End Sub

Private Sub GridCategoria_EnterCell()
     Call objCT.GridCategoria_EnterCell
End Sub

Private Sub GridCategoria_LeaveCell()
     Call objCT.GridCategoria_LeaveCell
End Sub

Private Sub GridCategoria_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridCategoria_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridCategoria_KeyPress(KeyAscii As Integer)
     Call objCT.GridCategoria_KeyPress(KeyAscii)
End Sub

Private Sub GridCategoria_LostFocus()
     Call objCT.GridCategoria_LostFocus
End Sub

Private Sub GridCategoria_RowColChange()
     Call objCT.GridCategoria_RowColChange
End Sub

Private Sub GridCategoria_Scroll()
     Call objCT.GridCategoria_Scroll
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub ComboCategoriaProduto_Change()
     Call objCT.ComboCategoriaProduto_Change
End Sub

Private Sub ComboCategoriaProduto_GotFocus()
     Call objCT.ComboCategoriaProduto_GotFocus
End Sub

Private Sub ComboCategoriaProduto_KeyPress(KeyAscii As Integer)
     Call objCT.ComboCategoriaProduto_KeyPress(KeyAscii)
End Sub

Private Sub ComboCategoriaProduto_Validate(Cancel As Boolean)
     Call objCT.ComboCategoriaProduto_Validate(Cancel)
End Sub

Private Sub ComboCategoriaProdutoItem_Change()
     Call objCT.ComboCategoriaProdutoItem_Change
End Sub

Private Sub ComboCategoriaProdutoItem_GotFocus()
     Call objCT.ComboCategoriaProdutoItem_GotFocus
End Sub

Private Sub ComboCategoriaProdutoItem_KeyPress(KeyAscii As Integer)
     Call objCT.ComboCategoriaProdutoItem_KeyPress(KeyAscii)
End Sub

Private Sub ComboCategoriaProdutoItem_Validate(Cancel As Boolean)
     Call objCT.ComboCategoriaProdutoItem_Validate(Cancel)
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Private Sub PercentMaisQuantCotacaoAnterior_Change()
     Call objCT.PercentMaisQuantCotacaoAnterior_Change
End Sub

Private Sub PercentMaisQuantCotacaoAnterior_Validate(Cancel As Boolean)
     Call objCT.PercentMaisQuantCotacaoAnterior_Validate(Cancel)
End Sub

Private Sub PercentMaisReceb_Change()
     Call objCT.PercentMaisReceb_Change
End Sub

Private Sub PercentMaisReceb_Validate(Cancel As Boolean)
     Call objCT.PercentMaisReceb_Validate(Cancel)
End Sub

Private Sub PercentMenosQuantCotacaoAnterior_Change()
     Call objCT.PercentMenosQuantCotacaoAnterior_Change
End Sub

Private Sub PercentMenosQuantCotacaoAnterior_Validate(Cancel As Boolean)
     Call objCT.PercentMenosQuantCotacaoAnterior_Validate(Cancel)
End Sub

Private Sub PercentMenosReceb_Change()
     Call objCT.PercentMenosReceb_Change
End Sub

Private Sub PercentMenosReceb_Validate(Cancel As Boolean)
     Call objCT.PercentMenosReceb_Validate(Cancel)
End Sub

Private Sub Label2_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label2(Index), Source, X, Y)
End Sub
Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2(Index), Button, Shift, X, Y)
End Sub
Private Sub Label26_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label26, Source, X, Y)
End Sub
Private Sub Label26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label26, Button, Shift, X, Y)
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
Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub
Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub
Private Sub LblUMEstoque_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblUMEstoque, Source, X, Y)
End Sub
Private Sub LblUMEstoque_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblUMEstoque, Button, Shift, X, Y)
End Sub
Private Sub NomeUMEstoque_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NomeUMEstoque, Source, X, Y)
End Sub
Private Sub NomeUMEstoque_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NomeUMEstoque, Button, Shift, X, Y)
End Sub
Private Sub LblUMCompras_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblUMCompras, Source, X, Y)
End Sub
Private Sub LblUMCompras_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblUMCompras, Button, Shift, X, Y)
End Sub
Private Sub NomeUMCompra_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NomeUMCompra, Source, X, Y)
End Sub
Private Sub NomeUMCompra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NomeUMCompra, Button, Shift, X, Y)
End Sub
Private Sub LblUMVendas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblUMVendas, Source, X, Y)
End Sub
Private Sub LblUMVendas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblUMVendas, Button, Shift, X, Y)
End Sub
Private Sub NomeUMVenda_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NomeUMVenda, Source, X, Y)
End Sub
Private Sub NomeUMVenda_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NomeUMVenda, Button, Shift, X, Y)
End Sub
Private Sub DescricaoClasseUM_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescricaoClasseUM, Source, X, Y)
End Sub
Private Sub DescricaoClasseUM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescricaoClasseUM, Button, Shift, X, Y)
End Sub
Private Sub LblClasseUM_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblClasseUM, Source, X, Y)
End Sub
Private Sub LblClasseUM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblClasseUM, Button, Shift, X, Y)
End Sub
Private Sub LabelContaProducao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContaProducao, Source, X, Y)
End Sub
Private Sub LabelContaProducao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContaProducao, Button, Shift, X, Y)
End Sub
Private Sub ContaContabilLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaContabilLabel, Source, X, Y)
End Sub
Private Sub ContaContabilLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaContabilLabel, Button, Shift, X, Y)
End Sub
Private Sub Label27_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label27, Source, X, Y)
End Sub
Private Sub Label27_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label27, Button, Shift, X, Y)
End Sub
Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub
Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub
Private Sub DescTipoProduto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescTipoProduto, Source, X, Y)
End Sub
Private Sub DescTipoProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescTipoProduto, Button, Shift, X, Y)
End Sub
Private Sub LblTipoProduto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblTipoProduto, Source, X, Y)
End Sub
Private Sub LblTipoProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblTipoProduto, Button, Shift, X, Y)
End Sub
Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub
Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub
Private Sub LabelNomeReduzido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeReduzido, Source, X, Y)
End Sub
Private Sub LabelNomeReduzido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeReduzido, Button, Shift, X, Y)
End Sub
Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub
Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub
Private Sub Label31_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label31, Source, X, Y)
End Sub
Private Sub Label31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label31, Button, Shift, X, Y)
End Sub
Private Sub Label30_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label30, Source, X, Y)
End Sub
Private Sub Label30_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label30, Button, Shift, X, Y)
End Sub
Private Sub DescSubst2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescSubst2, Source, X, Y)
End Sub
Private Sub DescSubst2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescSubst2, Button, Shift, X, Y)
End Sub
Private Sub DescSubst1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescSubst1, Source, X, Y)
End Sub
Private Sub DescSubst1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescSubst1, Button, Shift, X, Y)
End Sub
Private Sub LblSubst2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblSubst2, Source, X, Y)
End Sub
Private Sub LblSubst2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblSubst2, Button, Shift, X, Y)
End Sub
Private Sub LblSubst1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblSubst1, Source, X, Y)
End Sub
Private Sub LblSubst1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblSubst1, Button, Shift, X, Y)
End Sub
Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub
Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub
Private Sub QuantPedido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantPedido, Source, X, Y)
End Sub
Private Sub QuantPedido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantPedido, Button, Shift, X, Y)
End Sub
Private Sub Label32_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label32, Source, X, Y)
End Sub
Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label32, Button, Shift, X, Y)
End Sub
Private Sub Label23_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label23, Source, X, Y)
End Sub
Private Sub Label23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label23, Button, Shift, X, Y)
End Sub
Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub
Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub
Private Sub Label35_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label35, Source, X, Y)
End Sub
Private Sub Label35_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label35, Button, Shift, X, Y)
End Sub
Private Sub Label22_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label22, Source, X, Y)
End Sub
Private Sub Label22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label22, Button, Shift, X, Y)
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
Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub
Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub
Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub
Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub
Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub
Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub
Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub
Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub
Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub
Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub
Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub
Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub
Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub
Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub
Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub
Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub
Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub
Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub
Private Sub DescrUM_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescrUM, Source, X, Y)
End Sub
Private Sub DescrUM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescrUM, Button, Shift, X, Y)
End Sub
Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub
Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub
Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub
Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub
Private Sub ClasFiscIPI_Validate(Cancel As Boolean)
     Call objCT.ClasFiscIPI_Validate(Cancel)
End Sub

Private Sub LabelClassificacaoFiscal_Click()
     Call objCT.LabelClassificacaoFiscal_Click
End Sub

Private Sub BotaoTeste_Click()
     Call objCT.BotaoTeste_Click
End Sub

Private Sub SerieProx_Change()
     Call objCT.SerieProx_Change
End Sub

Private Sub SerieNum_Change()
     Call objCT.SerieNum_Change
End Sub

Private Sub SerieProx_Validate(Cancel As Boolean)
     Call objCT.SerieProx_Validate(Cancel)
End Sub

Private Sub SerieNum_Validate(Cancel As Boolean)
     Call objCT.SerieNum_Validate(Cancel)
End Sub

Private Sub BotaoProxNum_Click()
     Call objCT.BotaoProxNum_Click
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

Private Sub BotaoProxCodBarras_Click()
     Call objCT.BotaoProxCodBarras_Click
End Sub

Private Sub LabelGenero_Click()
     Call objCT.LabelGenero_Click
End Sub

Private Sub Genero_Change()
     Call objCT.Genero_Change
End Sub

Private Sub Genero_Validate(Cancel As Boolean)
     Call objCT.Genero_Validate(Cancel)
End Sub

Private Sub LabelISSQN_Click()
     Call objCT.LabelISSQN_Click
End Sub

Private Sub ISSQN_Change()
     Call objCT.ISSQN_Change
End Sub

Private Sub ISSQN_Validate(Cancel As Boolean)
     Call objCT.ISSQN_Validate(Cancel)
End Sub

Private Sub LabelCodServNFe_Click()
     Call objCT.LabelCodServNFe_Click
End Sub

Private Sub PrecoMaxConsumidor_Change()
     Call objCT.PrecoMaxConsumidor_Change
End Sub

Private Sub PrecoMaxConsumidor_Validate(Cancel As Boolean)
     Call objCT.PrecoMaxConsumidor_Validate(Cancel)
End Sub

Private Sub SiglaUMTrib_Click()
     Call objCT.SiglaUMTrib_Click
End Sub

Private Sub ProdutoEspecifico_Click()
     Call objCT.ProdutoEspecifico_Click
End Sub

Private Sub ExTIPI_Change()
     Call objCT.ExTIPI_Change
End Sub

Private Sub BotaoSRV_Click()
     Call objCT.BotaoSRV_Click
End Sub

Private Sub Fabricante_Change()
     Call objCT.Fabricante_Change
End Sub

Private Sub Fabricante_Click()
     Call objCT.Fabricante_Click
End Sub

Public Sub Fabricante_Validate(Cancel As Boolean)
     Call objCT.Fabricante_Validate(Cancel)
End Sub

Private Sub QdtePorEmb_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.QdtePorEmb_Change(objCT)
End Sub

Private Sub QdtePorEmb_GotFocus()
    Call objCT.gobjInfoUsu.gobjTelaUsu.QdtePorEmb_GotFocus(objCT)
End Sub

Private Sub QdtePorEmb_Validate(Cancel As Boolean)
    Call objCT.gobjInfoUsu.gobjTelaUsu.QdtePorEmb_Validate(objCT, Cancel)
End Sub
