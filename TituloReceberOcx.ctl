VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl TituloReceberOcx 
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10215
   KeyPreview      =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   10215
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5445
      Index           =   1
      Left            =   165
      TabIndex        =   0
      Top             =   735
      Width           =   9930
      Begin VB.Frame Frame6 
         Caption         =   "Reajuste"
         Height          =   570
         Left            =   3615
         TabIndex        =   112
         Top             =   4665
         Width           =   6255
         Begin VB.ComboBox ReajustePeriodicidade 
            Height          =   315
            ItemData        =   "TituloReceberOcx.ctx":0000
            Left            =   945
            List            =   "TituloReceberOcx.ctx":0016
            TabIndex        =   119
            Top             =   180
            Width           =   1350
         End
         Begin VB.ComboBox Moeda 
            Height          =   315
            Left            =   3060
            TabIndex        =   113
            Top             =   180
            Width           =   1110
         End
         Begin MSComCtl2.UpDown UpDownReajusteBase 
            Height          =   300
            Left            =   5820
            TabIndex        =   114
            TabStop         =   0   'False
            Top             =   165
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox ReajusteBase 
            Height          =   300
            Left            =   4725
            TabIndex        =   115
            Top             =   165
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Periodic.:"
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
            Left            =   75
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   120
            Top             =   255
            Width           =   825
         End
         Begin VB.Label LabelMoeda 
            AutoSize        =   -1  'True
            Caption         =   "Índice:"
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
            Left            =   2400
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   117
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label21 
            Caption         =   "Base:"
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
            Left            =   4200
            TabIndex        =   116
            Top             =   210
            Width           =   525
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tributos"
         Height          =   2490
         Left            =   45
         TabIndex        =   57
         Top             =   2745
         Width           =   3555
         Begin VB.Frame Frame2 
            Caption         =   "Retenções"
            Height          =   1440
            Left            =   135
            TabIndex        =   103
            Top             =   225
            Width           =   3300
            Begin MSMask.MaskEdBox ValorIRRF 
               Height          =   300
               Left            =   855
               TabIndex        =   104
               Top             =   255
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox PISRetido 
               Height          =   300
               Left            =   2295
               TabIndex        =   106
               Top             =   270
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox COFINSRetido 
               Height          =   300
               Left            =   855
               TabIndex        =   108
               Top             =   675
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox CSLLRetido 
               Height          =   300
               Left            =   2295
               TabIndex        =   110
               Top             =   675
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox ISSRetido 
               Height          =   300
               Left            =   855
               TabIndex        =   138
               Top             =   1065
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin VB.Label Label4 
               Caption         =   "ISS:"
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
               Left            =   450
               TabIndex        =   139
               Top             =   1080
               Width           =   405
            End
            Begin VB.Label Label3 
               Caption         =   "CSLL:"
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
               Left            =   1755
               TabIndex        =   111
               Top             =   705
               Width           =   480
            End
            Begin VB.Label Label2 
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
               Height          =   210
               Left            =   75
               TabIndex        =   109
               Top             =   705
               Width           =   765
            End
            Begin VB.Label Label1 
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
               Height          =   210
               Index           =   0
               Left            =   1860
               TabIndex        =   107
               Top             =   315
               Width           =   360
            End
            Begin VB.Label Label20 
               Caption         =   "IR:"
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
               Left            =   555
               TabIndex        =   105
               Top             =   285
               Width           =   270
            End
         End
         Begin VB.Frame SSFrame6 
            Caption         =   "INSS"
            Height          =   690
            Left            =   135
            TabIndex        =   59
            Top             =   1680
            Width           =   3300
            Begin VB.CheckBox INSSRetido 
               Caption         =   "Retido"
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
               Left            =   1830
               TabIndex        =   15
               Top             =   285
               Width           =   900
            End
            Begin MSMask.MaskEdBox ValorINSS 
               Height          =   300
               Left            =   840
               TabIndex        =   14
               Top             =   255
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label30 
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
               Height          =   210
               Left            =   270
               TabIndex        =   63
               Top             =   300
               Width           =   510
            End
         End
      End
      Begin VB.Frame SSFrame2 
         Caption         =   "Cabeçalho"
         Height          =   2505
         Left            =   45
         TabIndex        =   54
         Top             =   180
         Width           =   9855
         Begin VB.CommandButton BotaoProjetos 
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
            Height          =   315
            Left            =   3570
            TabIndex        =   13
            Top             =   1965
            Width           =   495
         End
         Begin VB.ComboBox Etapa 
            Height          =   315
            Left            =   7425
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1965
            Width           =   2295
         End
         Begin VB.CommandButton BotaoLimparFAT 
            Height          =   300
            Left            =   5205
            Picture         =   "TituloReceberOcx.ctx":004D
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Numeração Automática"
            Top             =   870
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.ComboBox Tipo 
            Height          =   315
            Left            =   1035
            TabIndex        =   5
            Top             =   870
            Width           =   2400
         End
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   4470
            TabIndex        =   2
            Top             =   300
            Width           =   1815
         End
         Begin MSMask.MaskEdBox NumTitulo 
            Height          =   300
            Left            =   4470
            TabIndex        =   6
            Top             =   870
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "99999999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1035
            TabIndex        =   1
            Top             =   300
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownEmissao 
            Height          =   300
            Left            =   8520
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   270
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   300
            Left            =   7425
            TabIndex        =   3
            Top             =   270
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Valor 
            Height          =   300
            Left            =   7425
            TabIndex        =   8
            Top             =   825
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Ccl 
            Height          =   300
            Left            =   7425
            TabIndex        =   10
            Top             =   1410
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Natureza 
            Height          =   300
            Left            =   1035
            TabIndex        =   9
            Top             =   1440
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Projeto 
            Height          =   285
            Left            =   1035
            TabIndex        =   11
            Top             =   1980
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Etapa:"
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
            Height          =   180
            Index           =   62
            Left            =   6780
            TabIndex        =   137
            Top             =   2025
            Width           =   570
         End
         Begin VB.Label LabelProjeto 
            AutoSize        =   -1  'True
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
            Height          =   180
            Left            =   270
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   136
            Top             =   2025
            Width           =   675
         End
         Begin VB.Label LabelNaturezaDesc 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2070
            TabIndex        =   122
            Top             =   1440
            Width           =   3210
         End
         Begin VB.Label LabelNatureza 
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
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   150
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   121
            Top             =   1455
            Width           =   840
         End
         Begin VB.Label CclLabel 
            AutoSize        =   -1  'True
            Caption         =   "Centro de Custo/Lucro:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   5355
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   118
            Top             =   1455
            Width           =   2010
         End
         Begin VB.Label NumeroFAT 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4470
            TabIndex        =   64
            Top             =   855
            Width           =   735
         End
         Begin VB.Label TipoDocumentoLabel 
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
            Left            =   540
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   65
            Top             =   900
            Width           =   450
         End
         Begin VB.Label Label17 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   6855
            TabIndex        =   66
            Top             =   855
            Width           =   510
         End
         Begin VB.Label Label16 
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
            Height          =   255
            Left            =   6615
            TabIndex        =   67
            Top             =   315
            Width           =   750
         End
         Begin VB.Label LabelFilial 
            AutoSize        =   -1  'True
            Caption         =   " Filial:"
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
            Left            =   3885
            TabIndex        =   68
            Top             =   345
            Width           =   525
         End
         Begin VB.Label NumeroLabel 
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
            Left            =   3675
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   69
            Top             =   915
            Width           =   720
         End
         Begin VB.Label ClienteEtiqueta 
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
            Left            =   330
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   70
            Top             =   345
            Width           =   660
         End
      End
      Begin VB.Frame SSFrame7 
         Caption         =   "Comissões na Emissão"
         Height          =   1935
         Left            =   3630
         TabIndex        =   62
         Top             =   2745
         Width           =   6240
         Begin VB.CommandButton BotaoVendedores 
            Caption         =   "Vendedores"
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
            Left            =   105
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   1485
            Width           =   1215
         End
         Begin MSMask.MaskEdBox VendedorEmissao 
            Height          =   255
            Left            =   630
            TabIndex        =   16
            Top             =   300
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercentualEmissao 
            Height          =   255
            Left            =   2085
            TabIndex        =   17
            Top             =   300
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
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
         Begin MSMask.MaskEdBox ValorBaseEmissao 
            Height          =   255
            Left            =   3345
            TabIndex        =   18
            Top             =   315
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   450
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
         Begin MSMask.MaskEdBox ValorEmissao 
            Height          =   255
            Left            =   4605
            TabIndex        =   19
            Top             =   300
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   450
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
         Begin MSFlexGridLib.MSFlexGrid GridComissoesEmissao 
            Height          =   1095
            Left            =   75
            TabIndex        =   20
            Top             =   270
            Width           =   5400
            _ExtentX        =   9525
            _ExtentY        =   1931
            _Version        =   393216
            Rows            =   4
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label TotalPercentualEmissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2520
            TabIndex        =   71
            Top             =   1470
            Width           =   945
         End
         Begin VB.Label TotalValorEmissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3840
            TabIndex        =   72
            Top             =   1470
            Width           =   1095
         End
         Begin VB.Label LabelTotaisEmissao 
            Caption         =   "Totais:"
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
            Left            =   1815
            TabIndex        =   73
            Top             =   1515
            Width           =   705
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4845
      Index           =   3
      Left            =   240
      TabIndex        =   33
      Top             =   750
      Visible         =   0   'False
      Width           =   9855
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4305
         TabIndex        =   140
         Tag             =   "1"
         Top             =   1485
         Width           =   870
      End
      Begin VB.CommandButton CTBBotaoImprimir 
         Caption         =   "Imprimir"
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
         Left            =   7830
         TabIndex        =   101
         Top             =   120
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6390
         Style           =   2  'Dropdown List
         TabIndex        =   100
         Top             =   960
         Width           =   2700
      End
      Begin VB.CommandButton CTBBotaoLimparGrid 
         Caption         =   "Limpar Grid"
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
         Left            =   6360
         TabIndex        =   99
         Top             =   120
         Width           =   1245
      End
      Begin VB.CommandButton CTBBotaoModeloPadrao 
         Caption         =   "Modelo Padrão"
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
         Left            =   6360
         TabIndex        =   98
         Top             =   435
         Width           =   2700
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4680
         TabIndex        =   42
         Top             =   1920
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
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
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   44
         Top             =   2565
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4230
         MaxLength       =   150
         TabIndex        =   43
         Top             =   2175
         Width           =   1770
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2985
         Left            =   6330
         TabIndex        =   46
         Top             =   1515
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   55
         Top             =   3450
         Width           =   5895
         Begin VB.Label CTBCclLabel 
            AutoSize        =   -1  'True
            Caption         =   "Centro de Custo:"
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
            TabIndex        =   74
            Top             =   660
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label CTBLabel7 
            AutoSize        =   -1  'True
            Caption         =   "Conta:"
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
            Left            =   1125
            TabIndex        =   75
            Top             =   300
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   76
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   77
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
      End
      Begin VB.CheckBox CTBLancAutomatico 
         Caption         =   "Recalcula Automaticamente"
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
         Left            =   3450
         TabIndex        =   37
         Top             =   960
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   38
         Top             =   1860
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDebito 
         Height          =   225
         Left            =   3435
         TabIndex        =   41
         Top             =   1890
         Width           =   1155
         _ExtentX        =   2037
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
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBCredito 
         Height          =   225
         Left            =   2280
         TabIndex        =   40
         Top             =   1830
         Width           =   1155
         _ExtentX        =   2037
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
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBCcl 
         Height          =   225
         Left            =   1545
         TabIndex        =   39
         Top             =   1875
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   10
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
      Begin MSComCtl2.UpDown CTBUpDown 
         Height          =   300
         Left            =   1650
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   525
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   570
         TabIndex        =   36
         Top             =   525
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBLote 
         Height          =   300
         Left            =   5580
         TabIndex        =   35
         Top             =   135
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDocumento 
         Height          =   300
         Left            =   3825
         TabIndex        =   34
         Top             =   120
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid CTBGridContabil 
         Height          =   1860
         Left            =   15
         TabIndex        =   45
         Top             =   1185
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSComctlLib.TreeView CTBTvwCcls 
         Height          =   2985
         Left            =   6330
         TabIndex        =   47
         Top             =   1515
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5265
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView CTBTvwContas 
         Height          =   2985
         Left            =   6330
         TabIndex        =   48
         Top             =   1515
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5265
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Label CTBLabel1 
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
         Left            =   6390
         TabIndex        =   102
         Top             =   750
         Width           =   690
      End
      Begin VB.Label CTBLabel21 
         Caption         =   "Origem:"
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
         TabIndex        =   78
         Top             =   165
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   79
         Top             =   120
         Width           =   1530
      End
      Begin VB.Label CTBLabel14 
         Caption         =   "Período:"
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
         Left            =   4230
         TabIndex        =   80
         Top             =   600
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   81
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   82
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label CTBLabel13 
         Caption         =   "Exercício:"
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
         Left            =   1995
         TabIndex        =   83
         Top             =   585
         Width           =   870
      End
      Begin VB.Label CTBLabel5 
         AutoSize        =   -1  'True
         Caption         =   "Lançamentos"
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
         TabIndex        =   84
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label CTBLabelHistoricos 
         Caption         =   "Históricos"
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
         Left            =   6345
         TabIndex        =   85
         Top             =   1275
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label CTBLabelContas 
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
         Height          =   255
         Left            =   6345
         TabIndex        =   86
         Top             =   1275
         Width           =   2340
      End
      Begin VB.Label CTBLabelCcl 
         Caption         =   "Centros de Custo / Lucro"
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
         Left            =   6345
         TabIndex        =   87
         Top             =   1275
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label CTBLabelTotais 
         Caption         =   "Totais:"
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
         Left            =   1800
         TabIndex        =   88
         Top             =   3045
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   89
         Top             =   3030
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   90
         Top             =   3030
         Width           =   1155
      End
      Begin VB.Label CTBLabel8 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   45
         TabIndex        =   91
         Top             =   555
         Width           =   480
      End
      Begin VB.Label CTBLabelDoc 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
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
         Left            =   2700
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   92
         Top             =   165
         Width           =   1035
      End
      Begin VB.Label CTBLabelLote 
         AutoSize        =   -1  'True
         Caption         =   "Lote:"
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
         Left            =   5100
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   93
         Top             =   165
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5430
      Index           =   2
      Left            =   165
      TabIndex        =   22
      Top             =   750
      Visible         =   0   'False
      Width           =   9945
      Begin VB.CommandButton BotaoVendedoresParc 
         Caption         =   "Vendedores"
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
         Left            =   8115
         Style           =   1  'Graphical
         TabIndex        =   135
         ToolTipText     =   "Lista de Diferenças nas Parcelas"
         Top             =   5055
         Width           =   1755
      End
      Begin VB.Frame SSFrame3 
         Caption         =   "Parcelas"
         Height          =   2760
         Left            =   0
         TabIndex        =   125
         Top             =   0
         Width           =   9915
         Begin VB.ComboBox CondicaoPagamento 
            Height          =   315
            Left            =   2970
            TabIndex        =   129
            Top             =   180
            Width           =   1950
         End
         Begin VB.CheckBox CobrancaAutomatica 
            Caption         =   "Calcula cobrança automaticamente"
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
            Left            =   5175
            TabIndex        =   128
            Top             =   210
            Value           =   1  'Checked
            Width           =   3360
         End
         Begin VB.CheckBox Previsao 
            Caption         =   "Previsão"
            Height          =   315
            Left            =   4590
            TabIndex        =   127
            Top             =   585
            Width           =   885
         End
         Begin VB.TextBox DescPrev 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   3150
            MaxLength       =   255
            TabIndex        =   126
            Top             =   975
            Width           =   4200
         End
         Begin MSMask.MaskEdBox DataVencimentoReal 
            Height          =   315
            Left            =   2160
            TabIndex        =   130
            Top             =   540
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorParcela 
            Height          =   315
            Left            =   3300
            TabIndex        =   131
            Top             =   540
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   315
            Left            =   1005
            TabIndex        =   132
            Top             =   510
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   1650
            Left            =   105
            TabIndex        =   133
            Top             =   690
            Width           =   9690
            _ExtentX        =   17092
            _ExtentY        =   2910
            _Version        =   393216
            Rows            =   50
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label CondPagtoLabel 
            Caption         =   "Condição de Pagamento:"
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
            Left            =   780
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   134
            Top             =   225
            Width           =   2175
         End
      End
      Begin VB.CommandButton BotaoEditarDif 
         Caption         =   "Editar Diferenças na Parcela"
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
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   124
         ToolTipText     =   "Lista de Diferenças nas Parcelas"
         Top             =   5055
         Width           =   3555
      End
      Begin VB.CommandButton BotaoDif 
         Caption         =   "Consultar Diferenças nas Parcelas"
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
         Left            =   4110
         Style           =   1  'Graphical
         TabIndex        =   123
         ToolTipText     =   "Lista de Diferenças nas Parcelas"
         Top             =   5055
         Width           =   3555
      End
      Begin VB.Frame SSFrame4 
         Caption         =   "Comissões"
         Height          =   2115
         Left            =   5025
         TabIndex        =   58
         Top             =   2835
         Width           =   4905
         Begin MSMask.MaskEdBox ValorComissao 
            Height          =   315
            Left            =   3075
            TabIndex        =   31
            Top             =   765
            Width           =   975
            _ExtentX        =   1720
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
         Begin MSMask.MaskEdBox ValorBase 
            Height          =   315
            Left            =   3135
            TabIndex        =   30
            Top             =   390
            Width           =   1050
            _ExtentX        =   1852
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
         Begin MSMask.MaskEdBox PercentualComissao 
            Height          =   315
            Left            =   1920
            TabIndex        =   29
            Top             =   615
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
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
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   315
            Left            =   645
            TabIndex        =   28
            Top             =   450
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridComissoes 
            Height          =   1125
            Left            =   30
            TabIndex        =   32
            Top             =   285
            Width           =   4845
            _ExtentX        =   8546
            _ExtentY        =   1984
            _Version        =   393216
            Rows            =   4
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label TotalPercentualComissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1350
            TabIndex        =   94
            Top             =   1755
            Width           =   945
         End
         Begin VB.Label TotalValorComissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2445
            TabIndex        =   95
            Top             =   1755
            Width           =   1155
         End
         Begin VB.Label LabelTotaisComissoes 
            Caption         =   "Totais:"
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
            Left            =   495
            TabIndex        =   96
            Top             =   1785
            Width           =   705
         End
      End
      Begin VB.Frame SSFrame1 
         Caption         =   "Descontos"
         Height          =   2100
         Left            =   0
         TabIndex        =   60
         Top             =   2835
         Width           =   4995
         Begin VB.ComboBox TipoDesconto 
            Height          =   315
            Left            =   585
            TabIndex        =   23
            Top             =   225
            Width           =   1605
         End
         Begin MSMask.MaskEdBox Percentual1 
            Height          =   315
            Left            =   3735
            TabIndex        =   26
            Top             =   585
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
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
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDesconto 
            Height          =   315
            Left            =   2115
            TabIndex        =   25
            Top             =   750
            Width           =   1005
            _ExtentX        =   1773
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
         Begin MSMask.MaskEdBox Data 
            Height          =   315
            Left            =   2505
            TabIndex        =   24
            Tag             =   "1"
            Top             =   240
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridDescontos 
            Height          =   1110
            Left            =   45
            TabIndex        =   27
            Top             =   285
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   1958
            _Version        =   393216
            Rows            =   4
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6900
      ScaleHeight     =   495
      ScaleWidth      =   3195
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   60
      Width           =   3255
      Begin VB.CommandButton BotaoDocOriginal 
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
         Left            =   60
         Picture         =   "TituloReceberOcx.ctx":057F
         Style           =   1  'Graphical
         TabIndex        =   97
         ToolTipText     =   "Consulta de documento original"
         Top             =   60
         Width           =   1065
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   390
         Left            =   2715
         Picture         =   "TituloReceberOcx.ctx":1489
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   390
         Left            =   2211
         Picture         =   "TituloReceberOcx.ctx":1607
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   390
         Left            =   1709
         Picture         =   "TituloReceberOcx.ctx":1B39
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   390
         Left            =   1207
         Picture         =   "TituloReceberOcx.ctx":1CC3
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5895
      Left            =   90
      TabIndex        =   61
      Top             =   375
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   10398
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Parcelas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contabilização"
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
Attribute VB_Name = "TituloReceberOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTTituloReceber
Attribute objCT.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
    Set objCT = New CTTituloReceber
    Set objCT.objUserControl = Me
End Sub

Private Sub BotaoDocOriginal_Click()
     Call objCT.BotaoDocOriginal_Click
End Sub

Private Sub BotaoEditarDif_Click()
     Call objCT.BotaoEditarDif_Click
End Sub

Private Sub BotaoLimparFAT_Click()
     Call objCT.BotaoLimparFAT_Click
End Sub

Private Sub CobrancaAutomatica_Click()
     Call objCT.CobrancaAutomatica_Click
End Sub

Private Sub Data_Change()
     Call objCT.Data_Change
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

Private Sub BotaoVendedores_Click()
     Call objCT.BotaoVendedores_Click
End Sub

Private Sub DataEmissao_GotFocus()
     Call objCT.DataEmissao_GotFocus
End Sub

Private Sub INSSRetido_Click()
     Call objCT.INSSRetido_Click
End Sub

Private Sub Moeda_Click()
     Call objCT.Moeda_Click
End Sub

Private Sub NumTitulo_GotFocus()
     Call objCT.NumTitulo_GotFocus
End Sub

Private Sub ClienteEtiqueta_Click()
     Call objCT.ClienteEtiqueta_Click
End Sub

Private Sub NumeroLabel_Click()
     Call objCT.NumeroLabel_Click
End Sub

Private Sub CondPagtoLabel_Click()
     Call objCT.CondPagtoLabel_Click
End Sub

Function Trata_Parametros(Optional objTituloReceber As ClassTituloReceber) As Long
     Trata_Parametros = objCT.Trata_Parametros(objTituloReceber)
End Function

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub CondicaoPagamento_Change()
     Call objCT.CondicaoPagamento_Change
End Sub

Private Sub CondicaoPagamento_Click()
     Call objCT.CondicaoPagamento_Click
End Sub

Private Sub CondicaoPagamento_Validate(bCancel As Boolean)
     Call objCT.CondicaoPagamento_Validate(bCancel)
End Sub

Private Sub DataEmissao_Change()
     Call objCT.DataEmissao_Change
End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
     Call objCT.DataEmissao_Validate(Cancel)
End Sub

Private Sub Opcao_GotFocus()
     Call objCT.Opcao_GotFocus
End Sub

Private Sub Percentual1_Change()
     Call objCT.Percentual1_Change
End Sub

Private Sub PercentualComissao_Change()
     Call objCT.PercentualComissao_Change
End Sub

Private Sub ReajustePeriodicidade_Click()
     Call objCT.ReajustePeriodicidade_Click
End Sub

Private Sub TipoDesconto_Change()
     Call objCT.TipoDesconto_Change
End Sub

Private Sub TipoDocumentoLabel_Click()
     Call objCT.TipoDocumentoLabel_Click
End Sub

Private Sub UpDownEmissao_DownClick()
     Call objCT.UpDownEmissao_DownClick
End Sub

Private Sub UpDownEmissao_UpClick()
     Call objCT.UpDownEmissao_UpClick
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub NumTitulo_Change()
     Call objCT.NumTitulo_Change
End Sub

Private Sub NumTitulo_Validate(Cancel As Boolean)
     Call objCT.NumTitulo_Validate(Cancel)
End Sub

Private Sub Cliente_Change()
     Call objCT.Cliente_Change
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)
     Call objCT.Cliente_Validate(Cancel)
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Private Sub GridParcelas_Click()
     Call objCT.GridParcelas_Click
End Sub

Private Sub GridParcelas_EnterCell()
     Call objCT.GridParcelas_EnterCell
End Sub

Private Sub GridParcelas_GotFocus()
     Call objCT.GridParcelas_GotFocus
End Sub

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridParcelas_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)
     Call objCT.GridParcelas_KeyPress(KeyAscii)
End Sub

Private Sub GridParcelas_LeaveCell()
     Call objCT.GridParcelas_LeaveCell
End Sub

Private Sub GridParcelas_Validate(Cancel As Boolean)
     Call objCT.GridParcelas_Validate(Cancel)
End Sub

Private Sub GridParcelas_RowColChange()
     Call objCT.GridParcelas_RowColChange
End Sub

Private Sub GridParcelas_Scroll()
     Call objCT.GridParcelas_Scroll
End Sub

Private Sub GridComissoes_Click()
     Call objCT.GridComissoes_Click
End Sub

Private Sub GridComissoes_EnterCell()
     Call objCT.GridComissoes_EnterCell
End Sub

Private Sub GridComissoes_GotFocus()
     Call objCT.GridComissoes_GotFocus
End Sub

Private Sub GridComissoes_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridComissoes_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridComissoes_KeyPress(KeyAscii As Integer)
     Call objCT.GridComissoes_KeyPress(KeyAscii)
End Sub

Private Sub GridComissoes_LeaveCell()
     Call objCT.GridComissoes_LeaveCell
End Sub

Private Sub GridComissoes_Validate(Cancel As Boolean)
     Call objCT.GridComissoes_Validate(Cancel)
End Sub

Private Sub GridComissoes_RowColChange()
     Call objCT.GridComissoes_RowColChange
End Sub

Private Sub GridComissoes_Scroll()
     Call objCT.GridComissoes_Scroll
End Sub

Private Sub GridDescontos_Click()
     Call objCT.GridDescontos_Click
End Sub

Private Sub GridDescontos_EnterCell()
     Call objCT.GridDescontos_EnterCell
End Sub

Private Sub GridDescontos_GotFocus()
     Call objCT.GridDescontos_GotFocus
End Sub

Private Sub GridDescontos_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridDescontos_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridDescontos_KeyPress(KeyAscii As Integer)
     Call objCT.GridDescontos_KeyPress(KeyAscii)
End Sub

Private Sub GridDescontos_LeaveCell()
     Call objCT.GridDescontos_LeaveCell
End Sub

Private Sub GridDescontos_Validate(Cancel As Boolean)
     Call objCT.GridDescontos_Validate(Cancel)
End Sub

Private Sub GridDescontos_RowColChange()
     Call objCT.GridDescontos_RowColChange
End Sub

Private Sub GridDescontos_Scroll()
     Call objCT.GridDescontos_Scroll
End Sub

Private Sub GridComissoesEmissao_Click()
     Call objCT.GridComissoesEmissao_Click
End Sub

Private Sub GridComissoesEmissao_EnterCell()
     Call objCT.GridComissoesEmissao_EnterCell
End Sub

Private Sub GridComissoesEmissao_GotFocus()
     Call objCT.GridComissoesEmissao_GotFocus
End Sub

Private Sub GridComissoesEmissao_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridComissoesEmissao_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridComissoesEmissao_KeyPress(KeyAscii As Integer)
     Call objCT.GridComissoesEmissao_KeyPress(KeyAscii)
End Sub

Private Sub GridComissoesEmissao_LeaveCell()
     Call objCT.GridComissoesEmissao_LeaveCell
End Sub

Private Sub GridComissoesEmissao_Validate(Cancel As Boolean)
     Call objCT.GridComissoesEmissao_Validate(Cancel)
End Sub

Private Sub GridComissoesEmissao_RowColChange()
     Call objCT.GridComissoesEmissao_RowColChange
End Sub

Private Sub GridComissoesEmissao_Scroll()
     Call objCT.GridComissoesEmissao_Scroll
End Sub

Private Sub Tipo_Change()
     Call objCT.Tipo_Change
End Sub

Private Sub Tipo_Click()
     Call objCT.Tipo_Click
End Sub

Private Sub Tipo_Validate(Cancel As Boolean)
     Call objCT.Tipo_Validate(Cancel)
End Sub

Private Sub Valor_Change()
     Call objCT.Valor_Change
End Sub

Private Sub Valor_Validate(Cancel As Boolean)
     Call objCT.Valor_Validate(Cancel)
End Sub

Private Sub ValorBase_Change()
     Call objCT.ValorBase_Change
End Sub

Private Sub ValorComissao_Change()
     Call objCT.ValorComissao_Change
End Sub

Private Sub ValorDesconto_Change()
     Call objCT.ValorDesconto_Change
End Sub

Private Sub ValorIRRF_Change()
     Call objCT.ValorIRRF_Change
End Sub

Private Sub ValorIRRF_Validate(Cancel As Boolean)
     Call objCT.ValorIRRF_Validate(Cancel)
End Sub

Private Sub ValorINSS_Change()
     Call objCT.ValorINSS_Change
End Sub

Private Sub ValorINSS_Validate(Cancel As Boolean)
     Call objCT.ValorINSS_Validate(Cancel)
End Sub

Private Sub ValorParcela_GotFocus()
     Call objCT.ValorParcela_GotFocus
End Sub

Private Sub ValorParcela_KeyPress(KeyAscii As Integer)
     Call objCT.ValorParcela_KeyPress(KeyAscii)
End Sub

Private Sub ValorParcela_Validate(Cancel As Boolean)
     Call objCT.ValorParcela_Validate(Cancel)
End Sub

Private Sub Vendedor_Change()
     Call objCT.Vendedor_Change
End Sub

Private Sub Vendedor_GotFocus()
     Call objCT.Vendedor_GotFocus
End Sub

Private Sub Vendedor_KeyPress(KeyAscii As Integer)
     Call objCT.Vendedor_KeyPress(KeyAscii)
End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)
     Call objCT.Vendedor_Validate(Cancel)
End Sub

Private Sub VendedorEmissao_GotFocus()
     Call objCT.VendedorEmissao_GotFocus
End Sub

Private Sub VendedorEmissao_KeyPress(KeyAscii As Integer)
     Call objCT.VendedorEmissao_KeyPress(KeyAscii)
End Sub

Private Sub VendedorEmissao_Validate(Cancel As Boolean)
     Call objCT.VendedorEmissao_Validate(Cancel)
End Sub

Private Sub PercentualComissao_GotFocus()
     Call objCT.PercentualComissao_GotFocus
End Sub

Private Sub PercentualComissao_KeyPress(KeyAscii As Integer)
     Call objCT.PercentualComissao_KeyPress(KeyAscii)
End Sub

Private Sub PercentualComissao_Validate(Cancel As Boolean)
     Call objCT.PercentualComissao_Validate(Cancel)
End Sub

Private Sub PercentualEmissao_GotFocus()
     Call objCT.PercentualEmissao_GotFocus
End Sub

Private Sub PercentualEmissao_KeyPress(KeyAscii As Integer)
     Call objCT.PercentualEmissao_KeyPress(KeyAscii)
End Sub

Private Sub PercentualEmissao_Validate(Cancel As Boolean)
     Call objCT.PercentualEmissao_Validate(Cancel)
End Sub

Private Sub ValorBase_GotFocus()
     Call objCT.ValorBase_GotFocus
End Sub

Private Sub ValorBase_KeyPress(KeyAscii As Integer)
     Call objCT.ValorBase_KeyPress(KeyAscii)
End Sub

Private Sub ValorBase_Validate(Cancel As Boolean)
     Call objCT.ValorBase_Validate(Cancel)
End Sub

Private Sub ValorBaseEmissao_GotFocus()
     Call objCT.ValorBaseEmissao_GotFocus
End Sub

Private Sub ValorBaseEmissao_KeyPress(KeyAscii As Integer)
     Call objCT.ValorBaseEmissao_KeyPress(KeyAscii)
End Sub

Private Sub ValorBaseEmissao_Validate(Cancel As Boolean)
     Call objCT.ValorBaseEmissao_Validate(Cancel)
End Sub

Private Sub ValorComissao_GotFocus()
     Call objCT.ValorComissao_GotFocus
End Sub

Private Sub ValorComissao_KeyPress(KeyAscii As Integer)
     Call objCT.ValorComissao_KeyPress(KeyAscii)
End Sub

Private Sub ValorComissao_Validate(Cancel As Boolean)
     Call objCT.ValorComissao_Validate(Cancel)
End Sub

Private Sub ValorEmissao_GotFocus()
     Call objCT.ValorEmissao_GotFocus
End Sub

Private Sub ValorEmissao_KeyPress(KeyAscii As Integer)
     Call objCT.ValorEmissao_KeyPress(KeyAscii)
End Sub

Private Sub ValorEmissao_Validate(Cancel As Boolean)
     Call objCT.ValorEmissao_Validate(Cancel)
End Sub

Private Sub TipoDesconto_GotFocus()
     Call objCT.TipoDesconto_GotFocus
End Sub

Private Sub TipoDesconto_KeyPress(KeyAscii As Integer)
     Call objCT.TipoDesconto_KeyPress(KeyAscii)
End Sub

Private Sub TipoDesconto_Validate(Cancel As Boolean)
     Call objCT.TipoDesconto_Validate(Cancel)
End Sub

Private Sub DataVencimento_GotFocus()
     Call objCT.DataVencimento_GotFocus
End Sub

Private Sub DataVencimento_KeyPress(KeyAscii As Integer)
     Call objCT.DataVencimento_KeyPress(KeyAscii)
End Sub

Private Sub DataVencimento_Validate(Cancel As Boolean)
     Call objCT.DataVencimento_Validate(Cancel)
End Sub

Private Sub Data_GotFocus()
     Call objCT.Data_GotFocus
End Sub

Private Sub Data_KeyPress(KeyAscii As Integer)
     Call objCT.Data_KeyPress(KeyAscii)
End Sub

Private Sub Data_Validate(Cancel As Boolean)
     Call objCT.Data_Validate(Cancel)
End Sub

Private Sub ValorDesconto_GotFocus()
     Call objCT.ValorDesconto_GotFocus
End Sub

Private Sub ValorDesconto_KeyPress(KeyAscii As Integer)
     Call objCT.ValorDesconto_KeyPress(KeyAscii)
End Sub

Private Sub ValorDesconto_Validate(Cancel As Boolean)
     Call objCT.ValorDesconto_Validate(Cancel)
End Sub

Private Sub Percentual1_GotFocus()
     Call objCT.Percentual1_GotFocus
End Sub

Private Sub Percentual1_KeyPress(KeyAscii As Integer)
     Call objCT.Percentual1_KeyPress(KeyAscii)
End Sub

Private Sub Percentual1_Validate(Cancel As Boolean)
     Call objCT.Percentual1_Validate(Cancel)
End Sub

Private Sub Filial_Change()
     Call objCT.Filial_Change
End Sub

Private Sub Filial_Click()
     Call objCT.Filial_Click
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
     Call objCT.Filial_Validate(Cancel)
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub CTBBotaoModeloPadrao_Click()
     Call objCT.CTBBotaoModeloPadrao_Click
End Sub

Private Sub CTBModelo_Click()
     Call objCT.CTBModelo_Click
End Sub

Private Sub CTBGridContabil_Click()
     Call objCT.CTBGridContabil_Click
End Sub

Private Sub CTBGridContabil_EnterCell()
     Call objCT.CTBGridContabil_EnterCell
End Sub

Private Sub CTBGridContabil_GotFocus()
     Call objCT.CTBGridContabil_GotFocus
End Sub

Private Sub CTBGridContabil_KeyPress(KeyAscii As Integer)
     Call objCT.CTBGridContabil_KeyPress(KeyAscii)
End Sub

Private Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.CTBGridContabil_KeyDown(KeyCode, Shift)
End Sub

Private Sub CTBGridContabil_LeaveCell()
     Call objCT.CTBGridContabil_LeaveCell
End Sub

Private Sub CTBGridContabil_Validate(Cancel As Boolean)
     Call objCT.CTBGridContabil_Validate(Cancel)
End Sub

Private Sub CTBGridContabil_RowColChange()
     Call objCT.CTBGridContabil_RowColChange
End Sub

Private Sub CTBGridContabil_Scroll()
     Call objCT.CTBGridContabil_Scroll
End Sub

Private Sub CTBConta_Change()
     Call objCT.CTBConta_Change
End Sub

Private Sub CTBConta_GotFocus()
     Call objCT.CTBConta_GotFocus
End Sub

Private Sub CTBConta_KeyPress(KeyAscii As Integer)
     Call objCT.CTBConta_KeyPress(KeyAscii)
End Sub

Private Sub CTBConta_Validate(Cancel As Boolean)
     Call objCT.CTBConta_Validate(Cancel)
End Sub

Private Sub CTBCcl_Change()
     Call objCT.CTBCcl_Change
End Sub

Private Sub CTBCcl_GotFocus()
     Call objCT.CTBCcl_GotFocus
End Sub

Private Sub CTBCcl_KeyPress(KeyAscii As Integer)
     Call objCT.CTBCcl_KeyPress(KeyAscii)
End Sub

Private Sub CTBCcl_Validate(Cancel As Boolean)
     Call objCT.CTBCcl_Validate(Cancel)
End Sub

Private Sub CTBCredito_Change()
     Call objCT.CTBCredito_Change
End Sub

Private Sub CTBCredito_GotFocus()
     Call objCT.CTBCredito_GotFocus
End Sub

Private Sub CTBCredito_KeyPress(KeyAscii As Integer)
     Call objCT.CTBCredito_KeyPress(KeyAscii)
End Sub

Private Sub CTBCredito_Validate(Cancel As Boolean)
     Call objCT.CTBCredito_Validate(Cancel)
End Sub

Private Sub CTBDebito_Change()
     Call objCT.CTBDebito_Change
End Sub

Private Sub CTBDebito_GotFocus()
     Call objCT.CTBDebito_GotFocus
End Sub

Private Sub CTBDebito_KeyPress(KeyAscii As Integer)
     Call objCT.CTBDebito_KeyPress(KeyAscii)
End Sub

Private Sub CTBDebito_Validate(Cancel As Boolean)
     Call objCT.CTBDebito_Validate(Cancel)
End Sub

Private Sub CTBSeqContraPartida_Change()
     Call objCT.CTBSeqContraPartida_Change
End Sub

Private Sub CTBSeqContraPartida_GotFocus()
     Call objCT.CTBSeqContraPartida_GotFocus
End Sub

Private Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)
     Call objCT.CTBSeqContraPartida_KeyPress(KeyAscii)
End Sub

Private Sub CTBSeqContraPartida_Validate(Cancel As Boolean)
     Call objCT.CTBSeqContraPartida_Validate(Cancel)
End Sub

Private Sub CTBHistorico_Change()
     Call objCT.CTBHistorico_Change
End Sub

Private Sub CTBHistorico_GotFocus()
     Call objCT.CTBHistorico_GotFocus
End Sub

Private Sub CTBHistorico_KeyPress(KeyAscii As Integer)
     Call objCT.CTBHistorico_KeyPress(KeyAscii)
End Sub

Private Sub CTBHistorico_Validate(Cancel As Boolean)
     Call objCT.CTBHistorico_Validate(Cancel)
End Sub

Private Sub CTBLancAutomatico_Click()
     Call objCT.CTBLancAutomatico_Click
End Sub

Private Sub CTBAglutina_Click()
     Call objCT.CTBAglutina_Click
End Sub

Private Sub CTBAglutina_GotFocus()
     Call objCT.CTBAglutina_GotFocus
End Sub

Private Sub CTBAglutina_KeyPress(KeyAscii As Integer)
     Call objCT.CTBAglutina_KeyPress(KeyAscii)
End Sub

Private Sub CTBAglutina_Validate(Cancel As Boolean)
     Call objCT.CTBAglutina_Validate(Cancel)
End Sub

Private Sub CTBGerencial_Click()
    Call objCT.CTBGerencial_Click
End Sub

Private Sub CTBGerencial_GotFocus()
    Call objCT.CTBGerencial_GotFocus
End Sub

Private Sub CTBGerencial_KeyPress(KeyAscii As Integer)
    Call objCT.CTBGerencial_KeyPress(KeyAscii)
End Sub

Private Sub CTBGerencial_Validate(Cancel As Boolean)
    Call objCT.CTBGerencial_Validate(Cancel)
End Sub

Private Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwContas_NodeClick(Node)
End Sub

Private Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwContas_Expand(Node)
End Sub

Private Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwCcls_NodeClick(Node)
End Sub

Private Sub CTBListHistoricos_DblClick()
     Call objCT.CTBListHistoricos_DblClick
End Sub

Private Sub CTBBotaoLimparGrid_Click()
     Call objCT.CTBBotaoLimparGrid_Click
End Sub

Private Sub CTBLote_Change()
     Call objCT.CTBLote_Change
End Sub

Private Sub CTBLote_GotFocus()
     Call objCT.CTBLote_GotFocus
End Sub

Private Sub CTBLote_Validate(Cancel As Boolean)
     Call objCT.CTBLote_Validate(Cancel)
End Sub

Private Sub CTBDataContabil_Change()
     Call objCT.CTBDataContabil_Change
End Sub

Private Sub CTBDataContabil_GotFocus()
     Call objCT.CTBDataContabil_GotFocus
End Sub

Private Sub CTBDataContabil_Validate(Cancel As Boolean)
     Call objCT.CTBDataContabil_Validate(Cancel)
End Sub

Private Sub CTBDocumento_Change()
     Call objCT.CTBDocumento_Change
End Sub

Private Sub CTBDocumento_GotFocus()
     Call objCT.CTBDocumento_GotFocus
End Sub

Private Sub CTBBotaoImprimir_Click()
     Call objCT.CTBBotaoImprimir_Click
End Sub

Private Sub CTBUpDown_DownClick()
     Call objCT.CTBUpDown_DownClick
End Sub

Private Sub CTBUpDown_UpClick()
     Call objCT.CTBUpDown_UpClick
End Sub

Private Sub CTBLabelDoc_Click()
     Call objCT.CTBLabelDoc_Click
End Sub

Private Sub CTBLabelLote_Click()
     Call objCT.CTBLabelLote_Click
End Sub

Private Sub Label30_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label30, Source, X, Y)
End Sub
Private Sub Label30_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label30, Button, Shift, X, Y)
End Sub
Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub
Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub
Private Sub NumeroFAT_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroFAT, Source, X, Y)
End Sub
Private Sub NumeroFAT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroFAT, Button, Shift, X, Y)
End Sub
Private Sub TipoDocumentoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoDocumentoLabel, Source, X, Y)
End Sub
Private Sub TipoDocumentoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoDocumentoLabel, Button, Shift, X, Y)
End Sub
Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub
Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub
Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub
Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub
Private Sub LabelFilial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilial, Source, X, Y)
End Sub
Private Sub LabelFilial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilial, Button, Shift, X, Y)
End Sub
Private Sub NumeroLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroLabel, Source, X, Y)
End Sub
Private Sub NumeroLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroLabel, Button, Shift, X, Y)
End Sub
Private Sub ClienteEtiqueta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ClienteEtiqueta, Source, X, Y)
End Sub
Private Sub ClienteEtiqueta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ClienteEtiqueta, Button, Shift, X, Y)
End Sub
Private Sub CondPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondPagtoLabel, Source, X, Y)
End Sub
Private Sub CondPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondPagtoLabel, Button, Shift, X, Y)
End Sub
Private Sub TotalPercentualEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalPercentualEmissao, Source, X, Y)
End Sub
Private Sub TotalPercentualEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalPercentualEmissao, Button, Shift, X, Y)
End Sub
Private Sub TotalValorEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalValorEmissao, Source, X, Y)
End Sub
Private Sub TotalValorEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalValorEmissao, Button, Shift, X, Y)
End Sub
Private Sub LabelTotaisEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotaisEmissao, Source, X, Y)
End Sub
Private Sub LabelTotaisEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotaisEmissao, Button, Shift, X, Y)
End Sub
Private Sub CTBCclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclLabel, Source, X, Y)
End Sub
Private Sub CTBCclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclLabel, Button, Shift, X, Y)
End Sub
Private Sub CTBLabel7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel7, Source, X, Y)
End Sub
Private Sub CTBLabel7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel7, Button, Shift, X, Y)
End Sub
Private Sub CTBContaDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBContaDescricao, Source, X, Y)
End Sub
Private Sub CTBContaDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBContaDescricao, Button, Shift, X, Y)
End Sub
Private Sub CTBCclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclDescricao, Source, X, Y)
End Sub
Private Sub CTBCclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclDescricao, Button, Shift, X, Y)
End Sub
Private Sub CTBLabel21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel21, Source, X, Y)
End Sub
Private Sub CTBLabel21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel21, Button, Shift, X, Y)
End Sub
Private Sub CTBOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBOrigem, Source, X, Y)
End Sub
Private Sub CTBOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBOrigem, Button, Shift, X, Y)
End Sub
Private Sub CTBLabel14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel14, Source, X, Y)
End Sub
Private Sub CTBLabel14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel14, Button, Shift, X, Y)
End Sub
Private Sub CTBPeriodo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBPeriodo, Source, X, Y)
End Sub
Private Sub CTBPeriodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBPeriodo, Button, Shift, X, Y)
End Sub
Private Sub CTBExercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBExercicio, Source, X, Y)
End Sub
Private Sub CTBExercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBExercicio, Button, Shift, X, Y)
End Sub
Private Sub CTBLabel13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel13, Source, X, Y)
End Sub
Private Sub CTBLabel13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel13, Button, Shift, X, Y)
End Sub
Private Sub CTBLabel5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel5, Source, X, Y)
End Sub
Private Sub CTBLabel5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel5, Button, Shift, X, Y)
End Sub
Private Sub CTBLabelHistoricos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelHistoricos, Source, X, Y)
End Sub
Private Sub CTBLabelHistoricos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelHistoricos, Button, Shift, X, Y)
End Sub
Private Sub CTBLabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelContas, Source, X, Y)
End Sub
Private Sub CTBLabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelContas, Button, Shift, X, Y)
End Sub
Private Sub CTBLabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelCcl, Source, X, Y)
End Sub
Private Sub CTBLabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelCcl, Button, Shift, X, Y)
End Sub
Private Sub CTBLabel1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel1, Source, X, Y)
End Sub
Private Sub CTBLabel1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel1, Button, Shift, X, Y)
End Sub
Private Sub CTBLabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelTotais, Source, X, Y)
End Sub
Private Sub CTBLabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelTotais, Button, Shift, X, Y)
End Sub
Private Sub CTBTotalDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalDebito, Source, X, Y)
End Sub
Private Sub CTBTotalDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalDebito, Button, Shift, X, Y)
End Sub
Private Sub CTBTotalCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalCredito, Source, X, Y)
End Sub
Private Sub CTBTotalCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalCredito, Button, Shift, X, Y)
End Sub
Private Sub CTBLabel8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel8, Source, X, Y)
End Sub
Private Sub CTBLabel8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel8, Button, Shift, X, Y)
End Sub
Private Sub CTBLabelDoc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelDoc, Source, X, Y)
End Sub
Private Sub CTBLabelDoc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelDoc, Button, Shift, X, Y)
End Sub
Private Sub CTBLabelLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelLote, Source, X, Y)
End Sub
Private Sub CTBLabelLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelLote, Button, Shift, X, Y)
End Sub
Private Sub TotalPercentualComissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalPercentualComissao, Source, X, Y)
End Sub
Private Sub TotalPercentualComissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalPercentualComissao, Button, Shift, X, Y)
End Sub
Private Sub TotalValorComissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalValorComissao, Source, X, Y)
End Sub
Private Sub TotalValorComissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalValorComissao, Button, Shift, X, Y)
End Sub
Private Sub LabelTotaisComissoes_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotaisComissoes, Source, X, Y)
End Sub
Private Sub LabelTotaisComissoes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotaisComissoes, Button, Shift, X, Y)
End Sub
Private Sub Opcao_BeforeClick(Cancel As Integer)
     Call objCT.Opcao_BeforeClick(Cancel)
End Sub

Private Sub PISRetido_Change()
     Call objCT.PISRetido_Change
End Sub

Private Sub PISRetido_Validate(Cancel As Boolean)
     Call objCT.PISRetido_Validate(Cancel)
End Sub

Private Sub COFINSRetido_Change()
     Call objCT.COFINSRetido_Change
End Sub

Private Sub COFINSRetido_Validate(Cancel As Boolean)
     Call objCT.COFINSRetido_Validate(Cancel)
End Sub

Private Sub CSLLRetido_Change()
     Call objCT.CSLLRetido_Change
End Sub

Private Sub CSLLRetido_Validate(Cancel As Boolean)
     Call objCT.CSLLRetido_Validate(Cancel)
End Sub

Private Sub DescPrev_Change()
     Call objCT.DescPrev_Change
End Sub

Private Sub DescPrev_GotFocus()
     Call objCT.DescPrev_GotFocus
End Sub

Private Sub DescPrev_KeyPress(KeyAscii As Integer)
     Call objCT.DescPrev_KeyPress(KeyAscii)
End Sub

Private Sub DescPrev_Validate(Cancel As Boolean)
     Call objCT.DescPrev_Validate(Cancel)
End Sub

Private Sub Previsao_Change()
     Call objCT.Previsao_Change
End Sub

Private Sub Previsao_GotFocus()
     Call objCT.Previsao_GotFocus
End Sub

Private Sub Previsao_KeyPress(KeyAscii As Integer)
     Call objCT.Previsao_KeyPress(KeyAscii)
End Sub

Private Sub Previsao_Validate(Cancel As Boolean)
     Call objCT.Previsao_Validate(Cancel)
End Sub

Private Sub Cliente_Preenche()
     Call objCT.Cliente_Preenche
End Sub

Private Sub CclLabel_Click()
     Call objCT.CclLabel_Click
End Sub

Private Sub CclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclLabel, Source, X, Y)
End Sub
Private Sub CclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclLabel, Button, Shift, X, Y)
End Sub
Private Sub Ccl_Validate(Cancel As Boolean)
     Call objCT.Ccl_Validate(Cancel)
End Sub

Private Sub ReajusteBase_Change()
     Call objCT.ReajusteBase_Change
End Sub

Private Sub ReajusteBase_Validate(Cancel As Boolean)
     Call objCT.ReajusteBase_Validate(Cancel)
End Sub

Private Sub UpDownReajusteBase_DownClick()
     Call objCT.UpDownReajusteBase_DownClick
End Sub

Private Sub UpDownReajusteBase_UpClick()
     Call objCT.UpDownReajusteBase_UpClick
End Sub

Private Sub BotaoDif_Click()
     Call objCT.BotaoDif_Click
End Sub

Private Sub BotaoVendedoresParc_Click()
     Call objCT.BotaoVendedoresParc_Click
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub
Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_UnLoad(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_UnLoad(Cancel)
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

Private Sub ISSRetido_Change()
     Call objCT.ISSRetido_Change
End Sub

Private Sub ISSRetido_Validate(Cancel As Boolean)
     Call objCT.ISSRetido_Validate(Cancel)
End Sub

Private Sub BotaoProjetos_Click()
    Call objCT.BotaoProjetos_Click
End Sub

Private Sub LabelProjeto_Click()
    Call objCT.LabelProjeto_Click
End Sub

Private Sub Projeto_Change()
     Call objCT.Projeto_Change
End Sub

Private Sub Projeto_GotFocus()
     Call objCT.Projeto_GotFocus
End Sub

Private Sub Projeto_Validate(Cancel As Boolean)
     Call objCT.Projeto_Validate(Cancel)
End Sub

Private Sub Natureza_Change()
    Call objCT.Natureza_Change
End Sub

Private Sub LabelNatureza_Click()
    Call objCT.LabelNatureza_Click
End Sub

Private Sub Natureza_Validate(Cancel As Boolean)
    Call objCT.Natureza_Validate(Cancel)
End Sub
