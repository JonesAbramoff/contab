VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpPrevisaoProdOcx 
   ClientHeight    =   4755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7515
   ScaleHeight     =   4755
   ScaleWidth      =   7515
   Begin VB.Frame SSFrame1 
      Caption         =   "Filtros"
      Height          =   3720
      Left            =   180
      TabIndex        =   37
      Top             =   885
      Width           =   6075
      Begin VB.Frame Frame1 
         Caption         =   "Pedidos de Compra"
         Height          =   2895
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   5655
         Begin VB.Frame Frame9 
            Caption         =   "Código"
            Height          =   705
            Left            =   120
            TabIndex        =   68
            Top             =   225
            Width           =   3840
            Begin MSMask.MaskEdBox CodPCDe 
               Height          =   300
               Left            =   450
               TabIndex        =   23
               Top             =   240
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodPCAte 
               Height          =   300
               Left            =   1800
               TabIndex        =   24
               Top             =   240
               Width           =   780
               _ExtentX        =   1376
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
               Left            =   105
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   70
               Top             =   285
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
               Left            =   1425
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   69
               Top             =   285
               Width           =   360
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Data de Envio"
            Height          =   705
            Left            =   120
            TabIndex        =   63
            Top             =   1920
            Width           =   3840
            Begin MSComCtl2.UpDown UpDownDataEnvioDe 
               Height          =   315
               Left            =   1680
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   180
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEnvioDe 
               Height          =   315
               Left            =   495
               TabIndex        =   27
               Top             =   195
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
               Left            =   3540
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   180
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEnvioAte 
               Height          =   315
               Left            =   2355
               TabIndex        =   28
               Top             =   195
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
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
               TabIndex        =   67
               Top             =   270
               Width           =   315
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
               Left            =   1995
               TabIndex        =   66
               Top             =   270
               Width           =   360
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Compradores"
            Height          =   705
            Left            =   120
            TabIndex        =   60
            Top             =   1065
            Width           =   3840
            Begin MSMask.MaskEdBox CompradorDe 
               Height          =   300
               Left            =   405
               TabIndex        =   25
               Top             =   255
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CompradorAte 
               Height          =   300
               Left            =   1845
               TabIndex        =   26
               Top             =   255
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
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
               Left            =   1485
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   62
               Top             =   330
               Width           =   360
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
               Left            =   75
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   61
               Top             =   315
               Width           =   315
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Fornecedores"
         Height          =   2895
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   5655
         Begin VB.Frame Frame7 
            Caption         =   "Nome Reduzido"
            Height          =   675
            Left            =   225
            TabIndex        =   57
            Top             =   1470
            Width           =   4995
            Begin MSMask.MaskEdBox NomeFornDe 
               Height          =   300
               Left            =   525
               TabIndex        =   20
               Top             =   240
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NomeFornAte 
               Height          =   300
               Left            =   3015
               TabIndex        =   21
               Top             =   240
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin VB.Label LabelNomeFornDe 
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
               Left            =   135
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   59
               Top             =   300
               Width           =   315
            End
            Begin VB.Label LabelNomeFornAte 
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
               Left            =   2565
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   58
               Top             =   315
               Width           =   360
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Código"
            Height          =   705
            Left            =   225
            TabIndex        =   54
            Top             =   390
            Width           =   4995
            Begin MSMask.MaskEdBox FornDe 
               Height          =   300
               Left            =   525
               TabIndex        =   18
               Top             =   270
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   8
               Mask            =   "########"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox FornAte 
               Height          =   300
               Left            =   3015
               TabIndex        =   19
               Top             =   270
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   8
               Mask            =   "########"
               PromptChar      =   " "
            End
            Begin VB.Label LabelCodigoFornDe 
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
               TabIndex        =   56
               Top             =   330
               Width           =   315
            End
            Begin VB.Label LabelCodigoFornAte 
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
               Left            =   2580
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   55
               Top             =   300
               Width           =   360
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Produtos"
         Height          =   2895
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   5655
         Begin VB.Frame Frame2 
            Caption         =   "Nome Reduzido"
            Height          =   675
            Left            =   195
            TabIndex        =   51
            Top             =   1305
            Width           =   5235
            Begin MSMask.MaskEdBox NomeProdDe 
               Height          =   300
               Left            =   555
               TabIndex        =   13
               Top             =   240
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NomeProdAte 
               Height          =   300
               Left            =   3060
               TabIndex        =   14
               Top             =   225
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin VB.Label LabelNomeProdDe 
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
               TabIndex        =   53
               Top             =   285
               Width           =   315
            End
            Begin VB.Label LabelNomeProdAte 
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
               Left            =   2625
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   52
               Top             =   270
               Width           =   360
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Código"
            Height          =   1050
            Left            =   195
            TabIndex        =   48
            Top             =   210
            Width           =   2325
            Begin MSMask.MaskEdBox CodigoProdDe 
               Height          =   300
               Left            =   720
               TabIndex        =   8
               Top             =   225
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   529
               _Version        =   393216
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodigoProdAte 
               Height          =   300
               Left            =   705
               TabIndex        =   9
               Top             =   622
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   529
               _Version        =   393216
               PromptChar      =   " "
            End
            Begin VB.Label LabelCodigoProdDe 
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
               Left            =   270
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   50
               Top             =   285
               Width           =   315
            End
            Begin VB.Label LabelCodigoProdAte 
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
               Left            =   255
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   49
               Top             =   690
               Width           =   360
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Natureza"
            Height          =   675
            Left            =   195
            TabIndex        =   45
            Top             =   2055
            Visible         =   0   'False
            Width           =   5250
            Begin VB.ComboBox ComboNaturezaAte 
               Height          =   315
               ItemData        =   "RelOpPrevisaoProdOcx.ctx":0000
               Left            =   3090
               List            =   "RelOpPrevisaoProdOcx.ctx":001C
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   225
               Visible         =   0   'False
               Width           =   1785
            End
            Begin VB.ComboBox ComboNaturezaDe 
               Height          =   315
               ItemData        =   "RelOpPrevisaoProdOcx.ctx":00A9
               Left            =   540
               List            =   "RelOpPrevisaoProdOcx.ctx":00C5
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   225
               Visible         =   0   'False
               Width           =   1665
            End
            Begin VB.Label Label5 
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
               Left            =   2625
               TabIndex        =   47
               Top             =   270
               Visible         =   0   'False
               Width           =   360
            End
            Begin VB.Label Label6 
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
               TabIndex        =   46
               Top             =   270
               Visible         =   0   'False
               Width           =   315
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Tipo"
            Height          =   1035
            Left            =   2625
            TabIndex        =   44
            Top             =   210
            Width           =   2775
            Begin VB.ComboBox ComboProduto 
               Height          =   315
               Left            =   1200
               TabIndex        =   12
               Top             =   600
               Width           =   885
            End
            Begin VB.OptionButton OptionTodosTiposProd 
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
               Height          =   225
               Left            =   75
               TabIndex        =   10
               Top             =   270
               Width           =   1890
            End
            Begin VB.OptionButton OptionUmTipoProd 
               Caption         =   "Apenas "
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
               Left            =   105
               TabIndex        =   11
               Top             =   600
               Width           =   1035
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Filial Empresa"
         Height          =   2895
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   5655
         Begin VB.Frame Frame4 
            Caption         =   "Código"
            Height          =   660
            Left            =   210
            TabIndex        =   41
            Top             =   480
            Width           =   3390
            Begin MSMask.MaskEdBox CodigoFilialDe 
               Height          =   300
               Left            =   540
               TabIndex        =   3
               Top             =   225
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodigoFilialAte 
               Height          =   300
               Left            =   2175
               TabIndex        =   4
               Top             =   225
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin VB.Label LabelCodigoAte 
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
               Left            =   1725
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   43
               Top             =   285
               Width           =   360
            End
            Begin VB.Label LabelCodigoDe 
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
               Left            =   210
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   42
               Top             =   285
               Width           =   315
            End
         End
         Begin VB.Frame FrameNome 
            Caption         =   "Nome"
            Height          =   675
            Left            =   195
            TabIndex        =   38
            Top             =   1380
            Width           =   5160
            Begin MSMask.MaskEdBox NomeFilialAte 
               Height          =   300
               Left            =   3075
               TabIndex        =   6
               Top             =   210
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NomeFilialDe 
               Height          =   300
               Left            =   555
               TabIndex        =   5
               Top             =   240
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin VB.Label LabelNomeDe 
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
               TabIndex        =   40
               Top             =   270
               Width           =   315
            End
            Begin VB.Label LabelNomeAte 
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
               Left            =   2625
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   39
               Top             =   270
               Width           =   360
            End
         End
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   3375
         Left            =   75
         TabIndex        =   71
         Top             =   240
         Width           =   5820
         _ExtentX        =   10266
         _ExtentY        =   5953
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Filiais Empresa"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Produtos"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Fornecedores"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Pedidos de Compra"
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
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPrevisaoProdOcx.ctx":0152
      Left            =   855
      List            =   "RelOpPrevisaoProdOcx.ctx":0154
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   2355
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
      Left            =   3330
      Picture         =   "RelOpPrevisaoProdOcx.ctx":0156
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   165
      Width           =   1635
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5220
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   195
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPrevisaoProdOcx.ctx":0258
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPrevisaoProdOcx.ctx":03D6
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPrevisaoProdOcx.ctx":0908
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPrevisaoProdOcx.ctx":0A92
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOrdenacao 
      Height          =   315
      ItemData        =   "RelOpPrevisaoProdOcx.ctx":0BEC
      Left            =   1575
      List            =   "RelOpPrevisaoProdOcx.ctx":0BF6
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   435
      Visible         =   0   'False
      Width           =   1635
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
      Left            =   210
      TabIndex        =   36
      Top             =   135
      Width           =   615
   End
   Begin VB.Label Label2 
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
      TabIndex        =   35
      Top             =   525
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "RelOpPrevisaoProdOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'RelOpPrevisaoProd
Const ORD_POR_CODIGO = 0
Const ORD_POR_DESCRICAO = 1

Private WithEvents objEventoCodigoFornDe As AdmEvento
Attribute objEventoCodigoFornDe.VB_VarHelpID = -1
Private WithEvents objEventoCodigoFornAte As AdmEvento
Attribute objEventoCodigoFornAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeFornDe As AdmEvento
Attribute objEventoNomeFornDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeFornAte As AdmEvento
Attribute objEventoNomeFornAte.VB_VarHelpID = -1
Private WithEvents objEventoCodigoFilialDe As AdmEvento
Attribute objEventoCodigoFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoCodigoFilialAte As AdmEvento
Attribute objEventoCodigoFilialAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialDe As AdmEvento
Attribute objEventoNomeFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialAte As AdmEvento
Attribute objEventoNomeFilialAte.VB_VarHelpID = -1
Private WithEvents objEventoCodProdDe As AdmEvento
Attribute objEventoCodProdDe.VB_VarHelpID = -1
Private WithEvents objEventoCodProdAte As AdmEvento
Attribute objEventoCodProdAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeProdDe As AdmEvento
Attribute objEventoNomeProdDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeProdAte As AdmEvento
Attribute objEventoNomeProdAte.VB_VarHelpID = -1
Private WithEvents objEventoCodPCDe As AdmEvento
Attribute objEventoCodPCDe.VB_VarHelpID = -1
Private WithEvents objEventoCodPCAte As AdmEvento
Attribute objEventoCodPCAte.VB_VarHelpID = -1
Private WithEvents objEventoCompradorDe As AdmEvento
Attribute objEventoCompradorDe.VB_VarHelpID = -1
Private WithEvents objEventoCompradorAte As AdmEvento
Attribute objEventoCompradorAte.VB_VarHelpID = -1
Private WithEvents objEventoFornDestino As AdmEvento
Attribute objEventoFornDestino.VB_VarHelpID = -1

Dim iAlterado As Integer
Dim iFrameAtual As Integer
Dim giTipoDestinoAtual As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio


Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 74046
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 74047
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 74046
        
        Case 74047
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171443)

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
    If lErro <> SUCESSO Then gError 74048
    
    'Call OptionTodosTipos_Click
    ComboOrdenacao.ListIndex = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    'FilialEmpresa.ListIndex = 0
    
    Exit Sub
    
Erro_Limpa_Tela_Rel:
    
    Select Case gErr
    
        Case 74048
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171444)

    End Select

    Exit Sub
   
End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel
   
End Sub

Private Sub CodigoFilialAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoFilialAte, iAlterado)
    
End Sub

Private Sub CodigoFilialDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoFilialDe, iAlterado)
    
End Sub


Private Sub CodigoProdDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_CodigoProdDe_Validate

    If Len(Trim(CodigoProdDe.ClipText)) > 0 Then
        
        lErro = CF("Produto_Formata", CodigoProdDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 74049
        
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 74050
        
        If lErro = 28030 Then gError 74051
        
    End If
    
    Exit Sub
    
Erro_CodigoProdDe_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 74049, 74050
        
        Case 74051
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171445)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub CodigoProdAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_CodigoProdAte_Validate

    If Len(Trim(CodigoProdAte.ClipText)) > 0 Then
        
        lErro = CF("Produto_Formata", CodigoProdAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 74052
        
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 74053
        
        If lErro = 28030 Then gError 74054
        
    End If
    
    Exit Sub
    
Erro_CodigoProdAte_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 74052, 74053
        
        Case 74054
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171446)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub CodPCAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodPCAte, iAlterado)
    
End Sub

Private Sub CodPCDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodPCDe, iAlterado)
    
End Sub

Private Sub ComboProduto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim bAchou As Boolean
Dim iIndice As Integer

On Error GoTo Erro_ComboProduto_Validate

    If Len(Trim(ComboProduto.Text)) = 0 Then Exit Sub
    
    For iIndice = 0 To ComboProduto.ListCount - 1
    
        If ComboProduto.Text = ComboProduto.List(iIndice) Then
            bAchou = True
            Exit For
        End If
    Next
    
    If bAchou = False Then gError 74055
    
    Exit Sub
    
Erro_ComboProduto_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 74055
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171447)
            
    End Select
    
    Exit Sub
    
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Form_Load
    
    iFrameAtual = 1
    
    Set objEventoCodigoFornDe = New AdmEvento
    Set objEventoCodigoFornAte = New AdmEvento
      
    Set objEventoFornDestino = New AdmEvento
    
    Set objEventoNomeFornDe = New AdmEvento
    Set objEventoNomeFornAte = New AdmEvento
    
    Set objEventoCodigoFilialDe = New AdmEvento
    Set objEventoCodigoFilialAte = New AdmEvento
    
    Set objEventoNomeFilialDe = New AdmEvento
    Set objEventoNomeFilialAte = New AdmEvento
    
    Set objEventoCodProdDe = New AdmEvento
    Set objEventoCodProdAte = New AdmEvento
    
    Set objEventoCodPCDe = New AdmEvento
    Set objEventoCodPCAte = New AdmEvento
    
    Set objEventoCompradorDe = New AdmEvento
    Set objEventoCompradorAte = New AdmEvento
    
    Set objEventoNomeProdDe = New AdmEvento
    Set objEventoNomeProdAte = New AdmEvento
    
    'Lê o Tipo de produto do BD
    lErro = CF("Cod_Nomes_Le", "TiposdeProduto", "TipoDeProduto", "Sigla", STRING_NOME_TABELA, colCodigoNome)
    If lErro <> SUCESSO Then gError 74058

    'Carrega a combo de Tipo de Produto
    For Each objCodigoNome In colCodigoNome
        ComboProduto.AddItem objCodigoNome.sNome
        ComboProduto.ItemData(ComboProduto.NewIndex) = objCodigoNome.iCodigo
    Next
    
    Set colCodigoNome = New AdmColCodigoNome
    
    OptionTodosTiposProd.Value = True
    
    'Inicializa as máscaras de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", CodigoProdDe)
    If lErro <> SUCESSO Then gError 74059

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", CodigoProdAte)
    If lErro <> SUCESSO Then gError 74060

    ComboOrdenacao.ListIndex = 1
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 74057, 74058, 74059, 74060, 74157
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171448)

    End Select

    Exit Sub

End Sub

Private Sub CompradorAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CompradorAte, iAlterado)
    
End Sub

Private Sub CompradorDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CompradorDe, iAlterado)
    
End Sub

Private Sub DataEnvioAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEnvioAte, iAlterado)
    
End Sub

Private Sub DataEnvioDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEnvioDe, iAlterado)

End Sub

Private Sub FornAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(FornAte, iAlterado)
    
End Sub

Private Sub FornAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornAte_Validate

    If Len(Trim(FornAte.Text)) > 0 Then

        'Lê o código informado
        objFornecedor.lCodigo = LCodigo_Extrai(FornAte.Text)
        
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then gError 74061
        
        'Se não encontrou o Fornecedor ==> erro
        If lErro = 12729 Then gError 74062
        
    End If

    Exit Sub

Erro_FornAte_Validate:

    Cancel = True

    Select Case gErr

        Case 74061

        Case 74062
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171449)

    End Select

    Exit Sub

End Sub

Private Sub FornDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(FornDe, iAlterado)
    
End Sub

Private Sub FornDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornDe_Validate

    If Len(Trim(FornDe.Text)) > 0 Then

        'Lê o código informado
        objFornecedor.lCodigo = LCodigo_Extrai(FornDe.Text)
        
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then gError 74063
        
        'Se não encontrou o Fornecedor ==> erro
        If lErro = 12729 Then gError 74064
        
    End If

    Exit Sub

Erro_FornDe_Validate:

    Cancel = True

    Select Case gErr

        Case 74063

        Case 74064
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171450)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigoFornAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelCodigoFornAte_Click
    
    If Len(Trim(FornAte.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.lCodigo = StrParaLong(FornAte.Text)
    End If
    
    'Chama Tela FornecedorsLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoCodigoFornAte)

   Exit Sub

Erro_LabelCodigoFornAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171451)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigoFornDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelCodigoFornDe_Click
    
    If Len(Trim(FornDe.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.lCodigo = StrParaLong(FornDe.Text)
    End If
    
    'Chama Tela FornecedorsLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoCodigoFornDe)

   Exit Sub

Erro_LabelCodigoFornDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171452)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEnvioDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 74142

    Exit Sub

Erro_UpDownDataEnvioDe_DownClick:

    Select Case gErr

        Case 74142
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 171453)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEnvioDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 74143

    Exit Sub

Erro_UpDownDataEnvioDe_UpClick:

    Select Case gErr

        Case 74143
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 171454)

    End Select

    Exit Sub

End Sub
Private Sub UpDownDataEnvioAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 74144

    Exit Sub

Erro_UpDownDataEnvioAte_DownClick:

    Select Case gErr

        Case 74144
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 171455)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEnvioAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 74145

    Exit Sub

Erro_UpDownDataEnvioAte_UpClick:

    Select Case gErr

        Case 74145
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 171456)

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
    If lErro <> SUCESSO Then gError 74147

    Exit Sub
                   
Erro_DataEnvioDe_Validate:

    Cancel = True

    Select Case gErr

        Case 74147
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171457)

    End Select

    Exit Sub

End Sub

Private Sub CompradorDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objComprador As New ClassComprador

On Error GoTo Erro_CompradorDe_Validate

    If Len(Trim(CompradorDe.Text)) > 0 Then

        lErro = CF("TP_Comprador_Le", CompradorDe, objComprador, 0)
        If lErro <> SUCESSO Then gError 74149
        
        CompradorDe.Text = objComprador.iCodigo
        
    End If

    Exit Sub

Erro_CompradorDe_Validate:

    Cancel = True


    Select Case gErr

        Case 74149

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171458)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171459)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171460)

    End Select

    Exit Sub

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171461)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171462)

    End Select

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

Private Sub CompradorAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objComprador As New ClassComprador

On Error GoTo Erro_CompradorAte_Validate

    If Len(Trim(CompradorAte.Text)) > 0 Then

        'Lê o código informado
        lErro = CF("TP_Comprador_Le", CompradorAte, objComprador, 0)
        If lErro <> SUCESSO Then gError 74148
        
        CompradorAte.Text = objComprador.iCodigo
        
    End If

    Exit Sub

Erro_CompradorAte_Validate:

    Cancel = True

    Select Case gErr

        Case 74148

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171463)

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
    If lErro <> SUCESSO Then gError 74146

    Exit Sub
                   
Erro_DataEnvioAte_Validate:

    Cancel = True

    Select Case gErr

        Case 74146
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171464)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigoProdDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_LabelCodigoProdDe_Click
    
    If Len(Trim(CodigoProdDe.Text)) > 0 Then
        'Preenche com o Produto da tela
        lErro = CF("Produto_Formata", CodigoProdDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 74065
        
        objProduto.sCodigo = sProdutoFormatado
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoCodProdDe)

   Exit Sub

Erro_LabelCodigoProdDe_Click:

    Select Case gErr

        Case 74065
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171465)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigoProdAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_LabelCodigoProdAte_Click
    
    If Len(Trim(CodigoProdAte.Text)) > 0 Then
        'Preenche com o Produto da tela
        lErro = CF("Produto_Formata", CodigoProdAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 74066
        
        objProduto.sCodigo = sProdutoFormatado
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoCodProdAte)

   Exit Sub

Erro_LabelCodigoProdAte_Click:

    Select Case gErr

        Case 74066
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171466)

    End Select

    Exit Sub

End Sub
Private Sub LabelNomeProdDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto

On Error GoTo Erro_LabelNomeProdDe_Click
    
    If Len(Trim(NomeProdDe.Text)) > 0 Then
        'Preenche com o Produto da tela
        objProduto.sNomeReduzido = NomeProdDe.Text
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoNomeProdDe)

   Exit Sub

Erro_LabelNomeProdDe_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171467)

    End Select

    Exit Sub
    
End Sub
Private Sub LabelNomeProdAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto

On Error GoTo Erro_LabelNomeProdAte_Click
    
    If Len(Trim(NomeProdAte.Text)) > 0 Then
        'Preenche com o Produto da tela
        objProduto.sNomeReduzido = NomeProdAte.Text
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoNomeProdAte)

   Exit Sub

Erro_LabelNomeProdAte_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171468)

    End Select

    Exit Sub
    
End Sub

Private Sub NomeFornAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_NomeFornAte_Validate

    'Verifica se o Nome do Fornecedor foi preenchido
    If Len(Trim(NomeFornAte.Text)) > 0 Then
    
        objFornecedor.sNomeReduzido = NomeFornAte.Text
        
        'Lê o Fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 74067
        If lErro = 6681 Then gError 74068

    End If
    
    Exit Sub
    
Erro_NomeFornAte_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 74067
        
        Case 74068
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171469)

    End Select
    
    Exit Sub
    
End Sub
Private Sub NomeFornDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_NomeFornDe_Validate

    'Verifica se o Nome do Fornecedor foi preenchido
    If Len(Trim(NomeFornDe.Text)) > 0 Then
    
        objFornecedor.sNomeReduzido = NomeFornDe.Text
        
        'Lê o Fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 74069
        If lErro = 6681 Then gError 74070

    End If
    
    Exit Sub
    
Erro_NomeFornDe_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 74069
        
        Case 74070
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171470)

    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoCodigoFilialAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    CodigoFilialAte.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodigoFilialDe_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    CodigoFilialDe.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodigoFornAte_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1
    
    FornAte.Text = CStr(objFornecedor.lCodigo)
    Call FornAte_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodigoFornDe_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1
    
    FornDe.Text = CStr(objFornecedor.lCodigo)
    Call FornDe_Validate(bSGECancelDummy)

    Me.Show

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171471)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171472)

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
        If lErro <> SUCESSO Then gError 74071

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = NomeFilialDe.Text Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 74072
        
        NomeFilialDe.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeFilialDe_Validate:

    Cancel = True

    Select Case gErr

        Case 74071

        Case 74072
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeFilialDe.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171473)

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
        If lErro <> SUCESSO Then gError 74073

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = NomeFilialAte.Text Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 74074

        NomeFilialAte.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 74073

        Case 74074
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeFilialAte.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171474)

    End Select

Exit Sub

End Sub

Private Sub CodigoFilialDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CodigoFilialDe_Validate

    If Len(Trim(CodigoFilialDe.Text)) > 0 Then

        'Lê o código informado
        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialDe.Text)
        
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 74075
        
        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 74076

    End If

    Exit Sub

Erro_CodigoFilialDe_Validate:

    Cancel = True


    Select Case gErr

        Case 74075

        Case 74076
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171475)

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
        If lErro <> SUCESSO And lErro <> 27378 Then gError 74077
        
        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 74078

    End If

    Exit Sub

Erro_CodigoFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 74077

        Case 74078
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171476)

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
        objFilialEmpresa.lCodEmpresa = StrParaLong(CodigoFilialDe.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoCodigoFilialDe)

   Exit Sub

Erro_LabelCodigoDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171477)

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
        objFilialEmpresa.lCodEmpresa = StrParaLong(CodigoFilialAte.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoCodigoFilialAte)

   Exit Sub

Erro_LabelCodigoAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171478)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeFornDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelNomeFornDe_Click
    
    If Len(Trim(NomeFornDe.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.sNomeReduzido = NomeFornDe.Text
    End If
    
    'Chama Tela FornecedorsLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoNomeFornDe)

   Exit Sub

Erro_LabelNomeFornDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171479)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeFornAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelNomeFornAte_Click
    
    If Len(Trim(NomeFornAte.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.sNomeReduzido = NomeFornAte.Text
    End If
    
    'Chama Tela FornecedorsLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoNomeFornAte)

   Exit Sub

Erro_LabelNomeFornAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171480)

    End Select

    Exit Sub

End Sub
Private Sub objEventoCodProdAte_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoCodProdAte_evSelecao

    Set objProduto = obj1

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 74079
    
    CodigoProdAte.Text = sProdutoMascarado

    Me.Show

    Exit Sub

Erro_objEventoCodProdAte_evSelecao:

    Select Case gErr
    
        Case 74079
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171481)
            
    End Select
    
    Exit Sub
    
End Sub
Private Sub objEventoCodProdDe_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoCodProdDe_evSelecao

    Set objProduto = obj1

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 74080
    
    CodigoProdDe.Text = sProdutoMascarado

    Me.Show

    Exit Sub

Erro_objEventoCodProdDe_evSelecao:

    Select Case gErr
    
        Case 74080
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171482)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoNomeFilialAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    NomeFilialAte.Text = objFilialEmpresa.sNome

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

Private Sub objEventoNomeFornDe_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1
    
    NomeFornDe.Text = objFornecedor.sNomeReduzido
    Call NomeFornDe_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoNomeFornAte_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1
    
    NomeFornAte.Text = objFornecedor.sNomeReduzido
    Call NomeFornAte_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeProdDe_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto

    Set objProduto = obj1
    
    NomeProdDe.Text = objProduto.sNomeReduzido

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoNomeProdAte_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto

    Set objProduto = obj1
    
    NomeProdAte.Text = objProduto.sNomeReduzido

    Me.Show

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 74081

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 74082

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 74083
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 74084
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 74081
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 74082, 74083, 74084
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171483)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 74085

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 74086

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 74085
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 74086

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171484)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 74087

    Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_CODIGO
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "TipoDestino", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FornCliDestino", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialDestino", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ProdutoCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataLimite", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FornecedorCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilFornCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpreaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "PedCompraCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemPedCompra", 1)
                
            Case ORD_POR_DESCRICAO

                Call gobjRelOpcoes.IncluirOrdenacao(1, "TipoDestino", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FornCliDestino", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialDestino", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ProdDesc", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ProdutoCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataLimite", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FornecedorCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilFornCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpreaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "PedCompraCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemPedCompra", 1)
                
            Case Else
                gError 74961

    End Select

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 74087, 74961

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171485)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sFornecedor_I As String
Dim sFornecedor_F As String
Dim sNomeForn_I As String
Dim sNomeForn_F As String
Dim sFilial_I As String
Dim sFilial_F As String
Dim sNomeFilial_I As String
Dim sNomeFilial_F As String
Dim sOrdenacaoPor As String
Dim sCheckTipo As String
Dim sFornecedorTipo As String
Dim sNomeProd_I As String
Dim sNomeProd_F As String
Dim sCodProd_I As String
Dim sCodProd_F As String
Dim sCheckTipoProd As String
Dim sProdutoTipo As String
Dim sNatureza_I As String
Dim sNatureza_F As String
Dim sOrd As String
Dim sCodPC_I As String
Dim sCodPC_F As String
Dim sComprador_I As String
Dim sComprador_F As String

On Error GoTo Erro_PreencherRelOp
 
    lErro = Formata_E_Critica_Parametros(sFornecedor_I, sFornecedor_F, sNomeForn_I, sNomeForn_F, sFilial_I, sFilial_F, sNomeFilial_I, sNomeFilial_F, sCodProd_I, sCodProd_F, sNomeProd_I, sNomeProd_F, sNatureza_I, sNatureza_F, sCheckTipo, sFornecedorTipo, sCheckTipoProd, sProdutoTipo, sCodPC_I, sCodPC_F, sComprador_I, sComprador_F)
    If lErro <> SUCESSO Then gError 74088

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 74089
         
    lErro = objRelOpcoes.IncluirParametro("NCODFORNINIC", sFornecedor_I)
    If lErro <> AD_BOOL_TRUE Then gError 74090
    
    lErro = objRelOpcoes.IncluirParametro("NCODPCINIC", sCodPC_I)
    If lErro <> AD_BOOL_TRUE Then gError 74168
    
    lErro = objRelOpcoes.IncluirParametro("NCODCOMPINIC", sComprador_I)
    If lErro <> AD_BOOL_TRUE Then gError 74169
    
    'Preenche a data envio inicial
    If Trim(DataEnvioDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DENVINIC", DataEnvioDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DENVINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 74170
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEFORNINIC", NomeFornDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 74091
    
    lErro = objRelOpcoes.IncluirParametro("NCODFILIALINIC", sFilial_I)
    If lErro <> AD_BOOL_TRUE Then gError 74092
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALINIC", NomeFilialDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 74093
    
    lErro = objRelOpcoes.IncluirParametro("NCODPRODINIC", sCodProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 74094
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEPRODINIC", NomeProdDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 74095
    
    lErro = objRelOpcoes.IncluirParametro("TNATUREZAPRODINIC", sNatureza_I)
    If lErro <> AD_BOOL_TRUE Then gError 74096
    
    lErro = objRelOpcoes.IncluirParametro("NCODFORNFIM", sFornecedor_F)
    If lErro <> AD_BOOL_TRUE Then gError 74097
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEFORNFIM", NomeFornAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 74098
        
    lErro = objRelOpcoes.IncluirParametro("NCODFILIALFIM", sFilial_F)
    If lErro <> AD_BOOL_TRUE Then gError 74099
    
    lErro = objRelOpcoes.IncluirParametro("NCODPCFIM", sCodPC_F)
    If lErro <> AD_BOOL_TRUE Then gError 74171
    
    lErro = objRelOpcoes.IncluirParametro("NCODCOMPFIM", sComprador_F)
    If lErro <> AD_BOOL_TRUE Then gError 74172
    
    'Preenche a data envio final
    If Trim(DataEnvioAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DENVFIM", DataEnvioAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DENVFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 74173
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALFIM", NomeFilialAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 74100
        
    lErro = objRelOpcoes.IncluirParametro("NCODPRODFIM", sCodProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 74101
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEPRODFIM", NomeProdAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 74102
    
    lErro = objRelOpcoes.IncluirParametro("TNATUREZAPRODFIM", sNatureza_F)
    If lErro <> AD_BOOL_TRUE Then gError 74103
        
    Select Case ComboOrdenacao.ListIndex
        
            Case ORD_POR_CODIGO
            
                sOrdenacaoPor = "Codigo"
                
            Case ORD_POR_DESCRICAO
                
                sOrdenacaoPor = "Descricao"
                
            Case Else
                gError 74104
                  
    End Select
        
    lErro = objRelOpcoes.IncluirParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> AD_BOOL_TRUE Then gError 74105

    sOrd = ComboOrdenacao.ListIndex
    lErro = objRelOpcoes.IncluirParametro("NORDENACAO", sOrd)
    If lErro <> AD_BOOL_TRUE Then gError 74106

    lErro = objRelOpcoes.IncluirParametro("TOPTIPO", sCheckTipo)
    If lErro <> AD_BOOL_TRUE Then gError 74108

    lErro = objRelOpcoes.IncluirParametro("TTIPOPROD", ComboProduto.Text)
    If lErro <> AD_BOOL_TRUE Then gError 74109

    lErro = objRelOpcoes.IncluirParametro("TOPTIPOP", sCheckTipoProd)
    If lErro <> AD_BOOL_TRUE Then gError 74110

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sFornecedor_I, sFornecedor_F, sNomeForn_I, sNomeForn_F, sFilial_I, sFilial_F, sNomeFilial_I, sNomeFilial_F, sNomeProd_I, sNomeProd_F, sCodProd_I, sCodProd_I, sNatureza_I, sNatureza_F, sFornecedorTipo, sCheckTipo, sOrdenacaoPor, sCheckTipoProd, sProdutoTipo, sCodPC_I, sCodPC_F, sComprador_I, sComprador_F)
    If lErro <> SUCESSO Then gError 74111

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 74088 To 74111
        
        Case 74168 To 74180
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171486)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sFornecedor_I As String, sFornecedor_F As String, sNomeForn_I As String, sNomeForn_F As String, sFilial_I As String, sFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sCodProd_I As String, sCodProd_F As String, sNomeProd_I As String, sNomeProd_F As String, sNatureza_I As String, sNatureza_F As String, sCheckTipo As String, sFornecedorTipo As String, sCheckProd As String, sProdutoTipo As String, sCodPC_I As String, sCodPC_F As String, sComprador_I As String, sComprador_F As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
       
    'critica Fornecedor Inicial e Final
    If FornDe.Text <> "" Then
        sFornecedor_I = CStr(FornDe.Text)
    Else
        sFornecedor_I = ""
    End If
    
    If FornAte.Text <> "" Then
        sFornecedor_F = CStr(FornAte.Text)
    Else
        sFornecedor_F = ""
    End If
            
    If sFornecedor_I <> "" And sFornecedor_F <> "" Then
        
        If CLng(sFornecedor_I) > CLng(sFornecedor_F) Then gError 74112
        
    End If
                
    'critica NomeFornecedor Inicial e Final
    If NomeFornDe.Text <> "" Then
        sNomeForn_I = NomeFornDe.Text
    Else
        sNomeForn_I = ""
    End If
    
    If NomeFornAte.Text <> "" Then
        sNomeForn_F = NomeFornAte.Text
    Else
        sNomeForn_F = ""
    End If
            
    If sNomeForn_I <> "" And sNomeForn_F <> "" Then
        
        If sNomeForn_I > sNomeForn_F Then gError 74113
        
    End If
    
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", CodigoProdDe.Text, sCodProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 74979

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sCodProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", CodigoProdAte.Text, sCodProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 74980

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sCodProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sCodProd_I > sCodProd_F Then gError 74114

    End If
    
    'critica Nome Produto Inicial e Final
    If NomeProdDe.Text <> "" Then
        sNomeProd_I = NomeProdDe.Text
    Else
        sNomeProd_I = ""
    End If
    
    If NomeProdAte.Text <> "" Then
        sNomeProd_F = NomeProdAte.Text
    Else
        sNomeProd_F = ""
    End If
            
    If sNomeProd_I <> "" And sNomeProd_F <> "" Then
        
        If sNomeProd_I > sNomeProd_F Then gError 74115
        
    End If
    
    'critica Filial Inicial e Final
    If CodigoFilialDe.Text <> "" Then
        sFilial_I = CStr(CodigoFilialDe.Text)
    Else
        sFilial_I = ""
    End If
    
    If CodigoFilialAte.Text <> "" Then
        sFilial_F = CStr(CodigoFilialAte.Text)
    Else
        sFilial_F = ""
    End If
            
    If sFilial_I <> "" And sFilial_F <> "" Then
        
        If CLng(sFilial_I) > CLng(sFilial_F) Then gError 74116
        
    End If
    
    'critica NomeFilial Inicial e Final
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
        
        If sNomeFilial_I > sNomeFilial_F Then gError 74117
        
    End If
    
    If ComboNaturezaDe.Text <> "" Then
        sNatureza_I = ComboNaturezaDe.Text
    Else
        sNatureza_I = ""
    End If
    
    If ComboNaturezaAte.Text <> "" Then
        sNatureza_F = ComboNaturezaAte.Text
    Else
        sNatureza_F = ""
    End If
            
    If sNatureza_I <> "" And sNatureza_F <> "" Then
        
        If sNatureza_I > sNatureza_F Then gError 74118
        
    End If
    
    'Se a opção para todos os tipos de Produto estiver selecionada
    If OptionTodosTiposProd.Value = True Then
        sCheckProd = "Todos"
        sProdutoTipo = ""
    
    'Se a opção para apenas um tipo estiver selecionada
    Else
    
        If ComboProduto.Text = "" Then gError 74120
        sCheckProd = "Um"
        sProdutoTipo = ComboProduto.Text
    
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

        If StrParaLong(sCodPC_I) > StrParaLong(sCodPC_F) Then gError 74181

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
        
        If StrParaInt(sComprador_I) > StrParaInt(sComprador_F) Then gError 74182
        
    End If
    
    'data de Envio inicial não pode ser maior que a final
    If Trim(DataEnvioDe.ClipText) <> "" And Trim(DataEnvioAte.ClipText) <> "" Then
    
         If CDate(DataEnvioDe.Text) > CDate(DataEnvioAte.Text) Then gError 74183
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 74112
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", gErr)
            FornDe.SetFocus
                
        Case 74113
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", gErr)
            NomeFornDe.SetFocus
            
        Case 74114
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            CodigoProdDe.SetFocus
            
        Case 74115
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            NomeProdDe.SetFocus
            
        Case 74116
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            CodigoFilialDe.SetFocus
            
        Case 74117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            NomeFilialDe.SetFocus
        
        Case 74118
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZA_INICIAL_MAIOR", gErr)
            ComboNaturezaDe.SetFocus
            
        Case 74119
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPO_FORNECEDOR_NAO_PREENCHIDO", gErr)
            
        Case 74120
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 74181
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PC_INICIAL_MAIOR", gErr)
            CodPCDe.SetFocus
        
        Case 74182
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_INICIAL_MAIOR", gErr)
            CompradorDe.SetFocus
            
        Case 74183
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAENVIO_INICIAL_MAIOR", gErr)
            DataEnvioDe.SetFocus
            
        Case 74979
            CodigoProdDe.SetFocus
            
        Case 74980
            CodigoProdAte.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171487)

    End Select

    Exit Function

End Function

                                                                        
Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sFornecedor_I As String, sFornecedor_F As String, sNomeForn_I As String, sNomeForn_F As String, sFilial_I As String, sFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sNomeProd_I As String, sNomeProd_F As String, sCodProd_I As String, sCodProd_F As String, sNatureza_I As String, sNatureza_F As String, sFornecedorTipo As String, sCheckTipo As String, sOrdenacaoPor As String, sCheckProd As String, sProdutoTipo As String, sCodPC_I As String, sCodPC_F As String, sComprador_I As String, sComprador_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Monta_Expressao_Selecao

   If sFornecedor_I <> "" Then sExpressao = "Forn >= " & Forprint_ConvLong(CLng(sFornecedor_I))

   If sFornecedor_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Forn <= " & Forprint_ConvLong(CLng(sFornecedor_F))

    End If
           
    If sCodPC_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "PC >= " & Forprint_ConvLong(StrParaLong(sCodPC_I))

    End If
   
    If sCodPC_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "PC <= " & Forprint_ConvLong(StrParaLong(sCodPC_F))

    End If
   
    If sComprador_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Comp >= " & Forprint_ConvInt(StrParaInt(sComprador_I))

    End If
   
    If sComprador_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Comp <= " & Forprint_ConvInt(StrParaInt(sComprador_F))

    End If
    
    If Trim(DataEnvioDe.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "EnvioInic"
        
    End If
    
    If Trim(DataEnvioAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "EnvioFim"

    End If
    
    If sNomeForn_I <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FornNomeInic"

    End If
    
    If sNomeForn_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FornNomeFim"

    End If
    
    If sNomeProd_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "ProdNomeInic"

    End If
    
    If sNomeProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "ProdNomeFim"

    End If
    
    If sCodProd_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Prod <= " & Forprint_ConvTexto(sCodProd_I)

    End If
    
    If sCodProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Prod <= " & Forprint_ConvTexto(sCodProd_F)

    End If
    
'    If sNatureza_I <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "NaturezaInic"
'
'    End If
'
'    If sNatureza_F <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "NaturezaFim"
'
'    End If
    
    If sFilial_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FECodInic"

    End If
    
    If sFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FECodFim"

    End If
           
    If sNomeFilial_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FENomeInic"

    End If
    
    If sNomeFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FENomeFim"

    End If
    
    'Se a opção para apenas um Tipo de Fornecedor estiver selecionada
    If sCheckTipo = "Um" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TForn = " & Forprint_ConvInt(StrParaInt(sFornecedorTipo))

    End If
    
    'Se a opção para apenas um Tipo de Produto estiver selecionada
    If sCheckProd = "Um" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TProd = " & Forprint_ConvTexto(sProdutoTipo)

    End If
    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case 74999
            'erro tratado na rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171488)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sTipoFornecedor As String, iTipo As Integer
Dim sOrdenacaoPor As String
Dim sTipoProduto As String
Dim iIndice As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 74121
   
    'pega Fornecedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFORNINIC", sParam)
    If lErro <> SUCESSO Then gError 74122
    
    FornDe.Text = sParam
    Call FornDe_Validate(bSGECancelDummy)
    
    'pega  Fornecedor final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFORNFIM", sParam)
    If lErro <> SUCESSO Then gError 74123
    
    FornAte.Text = sParam
    Call FornAte_Validate(bSGECancelDummy)
                                
    'pega Comprador inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODCOMPINIC", sParam)
    If lErro <> SUCESSO Then gError 74158
    
    CompradorDe.Text = sParam
    Call CompradorDe_Validate(bSGECancelDummy)
    
    'pega  Comprador final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODCOMPFIM", sParam)
    If lErro <> SUCESSO Then gError 74159
    
    CompradorAte.Text = sParam
    Call CompradorAte_Validate(bSGECancelDummy)
                                
    'pega Codigo inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODPCINIC", sParam)
    If lErro <> SUCESSO Then gError 74160
    
    CodPCDe.Text = sParam
    
    'pega  Codigo final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODPCFIM", sParam)
    If lErro <> SUCESSO Then gError 74161
    
    CodPCAte.Text = sParam
    
    'pega Nome do Fornecedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFORNINIC", sParam)
    If lErro <> SUCESSO Then gError 74124
    
    NomeFornDe.Text = sParam
    Call NomeFornDe_Validate(bSGECancelDummy)
    
    'pega  Nome do Fornecedor final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFORNFIM", sParam)
    If lErro <> SUCESSO Then gError 74125
    
    NomeFornAte.Text = sParam
    Call NomeFornAte_Validate(bSGECancelDummy)
                            
    'pega Nome do produto inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 74126
    
    NomeProdDe.Text = sParam
    
    'pega  Nome do produto final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 74127
    
    NomeProdAte.Text = sParam
    
    'pega codigo do produto inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 74128
    
    CodigoProdDe.PromptInclude = False
    CodigoProdDe.Text = sParam
    CodigoProdDe.PromptInclude = True
    
    Call CodigoProdDe_Validate(bSGECancelDummy)
    
    'pega  codigo do produto final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 74129
    
    CodigoProdAte.PromptInclude = False
    CodigoProdAte.Text = sParam
    CodigoProdAte.PromptInclude = True
    
    Call CodigoProdAte_Validate(bSGECancelDummy)
    
'    'pega natureza do produto inicial e exibe
'    lErro = objRelOpcoes.ObterParametro("TNATUREZAPRODINIC", sParam)
'    If lErro <> SUCESSO Then gError 74130
'
'    ComboNaturezaDe.Text = sParam
'
'    'pega natureza do produto final e exibe
'    lErro = objRelOpcoes.ObterParametro("TNATUREZAPRODFIM", sParam)
'    If lErro <> SUCESSO Then gError 74131
'
'    ComboNaturezaAte.Text = sParam
    
    'pega Filial inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 74132
    
    CodigoFilialDe.Text = sParam
    Call FornDe_Validate(bSGECancelDummy)
    
    'pega  Filial final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 74133
    
    CodigoFilialAte.Text = sParam
    Call CodigoFilialAte_Validate(bSGECancelDummy)
                                
    'pega Nome da Filial inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 74134
    
    NomeFilialDe.Text = sParam
    Call NomeFilialDe_Validate(bSGECancelDummy)
    
    'pega  Nome da Filial final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 74135
    
    NomeFilialAte.Text = sParam
    Call NomeFilialAte_Validate(bSGECancelDummy)
                
    'pega DataEnvio inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DENVINIC", sParam)
    If lErro <> SUCESSO Then gError 74162
    
    Call DateParaMasked(DataEnvioDe, CDate(sParam))
    
    'pega data de envio final e exibe
    lErro = objRelOpcoes.ObterParametro("DENVFIM", sParam)
    If lErro <> SUCESSO Then gError 74163

    Call DateParaMasked(DataEnvioAte, CDate(sParam))
        
    'pega  Tipo de produto  e exibe
    lErro = objRelOpcoes.ObterParametro("TOPTIPOP", sParam)
    If lErro <> SUCESSO Then gError 74138
                   
    If sParam = "Todos" Then
    
        Call OptionTodosTiposProd_Click
        
    Else
    
        'pega tipo de produto e exibe
        lErro = objRelOpcoes.ObterParametro("TTIPOPROD", sTipoProduto)
        If lErro <> SUCESSO Then gError 74139
                        
        OptionUmTipoProd.Value = True
        ComboProduto.Enabled = True
        ComboProduto.Text = sTipoFornecedor
        Call Combo_Seleciona(ComboProduto, iTipo)
        
    End If
    
    lErro = objRelOpcoes.ObterParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> SUCESSO Then gError 74140
    
    Select Case sOrdenacaoPor
        
            Case "Codigo"
            
                ComboOrdenacao.ListIndex = ORD_POR_CODIGO
            
            Case "Descricao"
            
                ComboOrdenacao.ListIndex = ORD_POR_DESCRICAO
                                            
            Case Else
                gError 74141
                  
    End Select
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 74121 To 74141
        
        Case 74158 To 74167
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171489)

    End Select

    Exit Function

End Function

Private Sub OptionTodosTiposProd_Click()

Dim lErro As Long

On Error GoTo Erro_OptionTodosTiposProd_Click

    ComboProduto.ListIndex = -1
    ComboProduto.Enabled = False
    OptionTodosTiposProd.Value = True
    
    Exit Sub

Erro_OptionTodosTiposProd_Click:

    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171490)

    End Select

    Exit Sub

End Sub

Private Sub OptionUmTipoProd_Click()

Dim lErro As Long

On Error GoTo Erro_OptionUmTipoProd_Click

    ComboProduto.ListIndex = -1
    ComboProduto.Enabled = True
    
    Exit Sub

Erro_OptionUmTipoProd_Click:

    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171491)

    End Select

    Exit Sub

End Sub


Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoFornDestino = Nothing
    
    Set objEventoCodigoFornDe = Nothing
    Set objEventoCodigoFornAte = Nothing
    
    Set objEventoNomeFornDe = Nothing
    Set objEventoNomeFornAte = Nothing
    
    Set objEventoCodigoFilialDe = Nothing
    Set objEventoCodigoFilialAte = Nothing
    
    Set objEventoNomeFilialDe = Nothing
    Set objEventoNomeFilialAte = Nothing
    
    Set objEventoCompradorDe = Nothing
    Set objEventoCompradorAte = Nothing
    
    Set objEventoNomeProdDe = Nothing
    Set objEventoNomeProdAte = Nothing
    
    Set objEventoCodPCDe = Nothing
    Set objEventoCodPCAte = Nothing
    
    Set objEventoCodProdDe = Nothing
    Set objEventoCodProdAte = Nothing
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_CADFORN
    Set Form_Load_Ocx = Me
    Caption = "Previsão de Entrega de Produtos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPrevisaoProd"
    
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


Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
            
            'Esconde o frame atual, mostra o novo
            Frame1(TabStrip1.SelectedItem.Index).Visible = True
            Frame1(iFrameAtual).Visible = False
            'Armazena novo valor de iFrameAtual
            iFrameAtual = TabStrip1.SelectedItem.Index

        End If
        
    
    Exit Sub

Erro_TabStrip1_Click:
    
    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171492)

    End Select

    Exit Sub

End Sub

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
        
        If Me.ActiveControl Is FornDe Then
            Call LabelCodigoFornDe_Click
        ElseIf Me.ActiveControl Is FornAte Then
            Call LabelCodigoFornAte_Click
        ElseIf Me.ActiveControl Is NomeFornDe Then
            Call LabelNomeFornDe_Click
        ElseIf Me.ActiveControl Is NomeFornAte Then
            Call LabelNomeFornAte_Click
        ElseIf Me.ActiveControl Is CompradorDe Then
            Call LabelCompradorDe_Click
        ElseIf Me.ActiveControl Is CompradorAte Then
            Call LabelCompradorAte_Click
        ElseIf Me.ActiveControl Is CodPCDe Then
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
        ElseIf Me.ActiveControl Is CodigoProdDe Then
            Call LabelCodigoProdDe_Click
        ElseIf Me.ActiveControl Is CodigoProdAte Then
            Call LabelCodigoProdAte_Click
        ElseIf Me.ActiveControl Is NomeProdDe Then
            Call LabelNomeProdDe_Click
        ElseIf Me.ActiveControl Is NomeProdAte Then
            Call LabelNomeProdAte_Click
        End If
    
    End If

End Sub



Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoAte, Source, X, Y)
End Sub

Private Sub LabelCodigoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoDe, Source, X, Y)
End Sub

Private Sub LabelCodigoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoDe, Button, Shift, X, Y)
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

Private Sub LabelNomeProdDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeProdDe, Source, X, Y)
End Sub

Private Sub LabelNomeProdDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeProdDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeProdAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeProdAte, Source, X, Y)
End Sub

Private Sub LabelNomeProdAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeProdAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoProdDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoProdDe, Source, X, Y)
End Sub

Private Sub LabelCodigoProdDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoProdDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoProdAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoProdAte, Source, X, Y)
End Sub

Private Sub LabelCodigoProdAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoProdAte, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeFornDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeFornDe, Source, X, Y)
End Sub

Private Sub LabelNomeFornDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeFornDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeFornAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeFornAte, Source, X, Y)
End Sub

Private Sub LabelNomeFornAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeFornAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoFornDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoFornDe, Source, X, Y)
End Sub

Private Sub LabelCodigoFornDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoFornDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoFornAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoFornAte, Source, X, Y)
End Sub

Private Sub LabelCodigoFornAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoFornAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodPCDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodPCDe, Source, X, Y)
End Sub

Private Sub LabelCodPCDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodPCDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodPCAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodPCAte, Source, X, Y)
End Sub

Private Sub LabelCodPCAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodPCAte, Button, Shift, X, Y)
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
