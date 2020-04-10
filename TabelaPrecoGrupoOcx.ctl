VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl TabelaPrecoGrupoOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   4575
      Index           =   1
      Left            =   135
      TabIndex        =   34
      Top             =   1335
      Width           =   9210
      Begin VB.Frame Frame9 
         Caption         =   "Padrão para Geração"
         Height          =   1170
         Index           =   3
         Left            =   30
         TabIndex        =   41
         Top             =   3405
         Width           =   9150
         Begin VB.Frame Frame9 
            Caption         =   "Preço Novo"
            Height          =   915
            Index           =   6
            Left            =   465
            TabIndex        =   59
            Top             =   165
            Width           =   4650
            Begin VB.OptionButton OptPreco 
               Caption         =   "Reajustar preço antigo em "
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
               Index           =   1
               Left            =   150
               TabIndex        =   19
               Top             =   585
               Width           =   2700
            End
            Begin VB.OptionButton OptPreco 
               Caption         =   "Fixar novo preço em"
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
               Index           =   0
               Left            =   150
               TabIndex        =   17
               Top             =   225
               Value           =   -1  'True
               Width           =   2160
            End
            Begin MSMask.MaskEdBox PrecoNovoRS 
               Height          =   315
               Left            =   2325
               TabIndex        =   18
               Top             =   165
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox PrecoNovoPerc 
               Height          =   315
               Left            =   2850
               TabIndex        =   20
               Top             =   555
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   8
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "%"
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
               Index           =   8
               Left            =   3870
               TabIndex        =   61
               Top             =   630
               Width           =   150
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "reais"
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
               Index           =   6
               Left            =   3930
               TabIndex        =   60
               Top             =   225
               Width           =   420
            End
         End
         Begin MSMask.MaskEdBox TextoGrade 
            Height          =   315
            Left            =   6870
            TabIndex        =   21
            ToolTipText     =   "Utilizado da descrição do produto de grade"
            Top             =   285
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Texto Grade:"
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
            Index           =   7
            Left            =   5640
            TabIndex        =   42
            Top             =   330
            Width           =   1125
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Filtros por Produto"
         Height          =   3375
         Index           =   1
         Left            =   30
         TabIndex        =   35
         Top             =   30
         Width           =   9150
         Begin VB.Frame Frame9 
            Caption         =   "Busca por parte de texto dentro dos campos"
            Height          =   2100
            Index           =   5
            Left            =   4935
            TabIndex        =   46
            Top             =   1185
            Width           =   4140
            Begin VB.TextBox ModeloLike 
               Height          =   312
               Left            =   1935
               MaxLength       =   20
               TabIndex        =   16
               ToolTipText     =   "a% = Começa com a, %a = Termina com a e %a% = Possui a em qualquer parte"
               Top             =   1710
               Width           =   2145
            End
            Begin VB.TextBox ReferenciaLike 
               Height          =   312
               Left            =   1935
               MaxLength       =   20
               TabIndex        =   15
               ToolTipText     =   "a% = Começa com a, %a = Termina com a e %a% = Possui a em qualquer parte"
               Top             =   1335
               Width           =   2145
            End
            Begin VB.TextBox NomeRedLike 
               Height          =   312
               Left            =   1935
               MaxLength       =   20
               TabIndex        =   14
               ToolTipText     =   "a% = Começa com a, %a = Termina com a e %a% = Possui a em qualquer parte"
               Top             =   975
               Width           =   2145
            End
            Begin VB.TextBox DescricaoLike 
               Height          =   312
               Left            =   1935
               MaxLength       =   20
               TabIndex        =   13
               ToolTipText     =   "a% = Começa com a, %a = Termina com a e %a% = Possui a em qualquer parte"
               Top             =   600
               Width           =   2145
            End
            Begin VB.TextBox CodigoLike 
               Height          =   312
               Left            =   1935
               MaxLength       =   20
               TabIndex        =   12
               ToolTipText     =   "a% = Começa com a, %a = Termina com a e %a% = Possui a em qualquer parte"
               Top             =   240
               Width           =   2145
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Referência LIKE"
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
               Index           =   1
               Left            =   405
               TabIndex        =   51
               Top             =   1410
               Width           =   1395
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Código LIKE"
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
               Left            =   765
               TabIndex        =   50
               Top             =   300
               Width           =   1050
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Descrição LIKE"
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
               Left            =   495
               TabIndex        =   49
               Top             =   675
               Width           =   1320
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nome Reduzido LIKE"
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
               Index           =   2
               Left            =   30
               TabIndex        =   48
               Top             =   1050
               Width           =   1815
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Modelo LIKE"
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
               Index           =   0
               Left            =   735
               TabIndex        =   47
               Top             =   1770
               Width           =   1095
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Categorias"
            Height          =   2100
            Index           =   2
            Left            =   75
            TabIndex        =   36
            Top             =   1185
            Width           =   4785
            Begin VB.ComboBox ComboCategoriaProdutoItem 
               Height          =   315
               Left            =   2400
               TabIndex        =   45
               Top             =   1620
               Width           =   2190
            End
            Begin VB.ComboBox ComboCategoriaProduto 
               Height          =   315
               Left            =   255
               TabIndex        =   44
               Top             =   1620
               Width           =   1590
            End
            Begin MSFlexGridLib.MSFlexGrid GridCategoria 
               Height          =   1830
               Left            =   75
               TabIndex        =   11
               Top             =   210
               Width           =   4620
               _ExtentX        =   8149
               _ExtentY        =   3228
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
         Begin VB.CheckBox AnaliticoGrade 
            Caption         =   "Analíticos sem Grade"
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
            Left            =   1545
            TabIndex        =   6
            Top             =   225
            Value           =   1  'Checked
            Width           =   2280
         End
         Begin VB.CheckBox AnaliticoGrade 
            Caption         =   "Grades e Kits de Venda"
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
            Left            =   4125
            TabIndex        =   7
            Top             =   225
            Value           =   1  'Checked
            Width           =   2490
         End
         Begin VB.CheckBox AnaliticoGrade 
            Caption         =   "Analíticos com Grade"
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
            Index           =   2
            Left            =   6825
            TabIndex        =   8
            Top             =   225
            Width           =   2265
         End
         Begin MSMask.MaskEdBox ProdutoPai 
            Height          =   315
            Left            =   1545
            TabIndex        =   10
            Top             =   855
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TipoProduto 
            Height          =   315
            Left            =   1545
            TabIndex        =   9
            Top             =   465
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label LabelProdutoPai 
            Caption         =   "Produto Pai:"
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
            Height          =   255
            Left            =   405
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   40
            Top             =   885
            Width           =   1215
         End
         Begin VB.Label DescProdPai 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4140
            TabIndex        =   39
            Top             =   855
            Width           =   4860
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1050
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   38
            Top             =   525
            Width           =   450
         End
         Begin VB.Label DescTipoProduto 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2205
            TabIndex        =   37
            Top             =   465
            Width           =   6795
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   4575
      Index           =   2
      Left            =   150
      TabIndex        =   33
      Top             =   1305
      Visible         =   0   'False
      Width           =   9195
      Begin VB.CommandButton BotaoProduto 
         Caption         =   "Consultar Produto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   30
         TabIndex        =   25
         ToolTipText     =   "Abre o Cadastro do Produto Selecionado"
         Top             =   4140
         Width           =   1665
      End
      Begin VB.CommandButton BotaoGravar 
         Caption         =   "Gravar Preços"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7455
         TabIndex        =   26
         ToolTipText     =   "Grava os Preços que Foram Preenchidos"
         Top             =   4140
         Width           =   1665
      End
      Begin VB.Frame Frame9 
         Caption         =   "Produtos"
         Height          =   4125
         Index           =   4
         Left            =   45
         TabIndex        =   43
         Top             =   0
         Width           =   9090
         Begin MSMask.MaskEdBox ProdPrecoNovo 
            Height          =   225
            Left            =   7410
            TabIndex        =   58
            Top             =   765
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
            Format          =   "#,##0.00####"
            PromptChar      =   " "
         End
         Begin VB.TextBox ProdUM 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   5535
            TabIndex        =   56
            Top             =   885
            Width           =   615
         End
         Begin VB.TextBox ProdCodigo 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   510
            TabIndex        =   55
            Top             =   885
            Width           =   2295
         End
         Begin VB.TextBox ProdDesc 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   2835
            TabIndex        =   54
            Top             =   885
            Width           =   2640
         End
         Begin MSFlexGridLib.MSFlexGrid GridProdutos 
            Height          =   660
            Left            =   30
            TabIndex        =   24
            Top             =   240
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   1164
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox ProdPrecoAntigo 
            Height          =   225
            Left            =   6195
            TabIndex        =   57
            Top             =   930
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
            Format          =   "#,##0.00####"
            PromptChar      =   " "
         End
         Begin VB.Label Label41 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   53
            Top             =   3765
            Width           =   915
         End
         Begin VB.Label Descricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1140
            TabIndex        =   52
            Top             =   3735
            Width           =   7890
         End
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Tabela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Index           =   0
      Left            =   90
      TabIndex        =   28
      Top             =   0
      Width           =   8130
      Begin VB.ComboBox Tabela 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   180
         Width           =   1395
      End
      Begin VB.CommandButton BotaoEditarTabela 
         Caption         =   "Editar"
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
         Left            =   5730
         TabIndex        =   4
         Top             =   525
         Width           =   1080
      End
      Begin VB.CommandButton BotaoExcluirTabela 
         Caption         =   "Excluir"
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
         Left            =   6945
         TabIndex        =   5
         Top             =   525
         Width           =   1080
      End
      Begin VB.CommandButton BotaoCriarTabela 
         Caption         =   "Criar"
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
         Left            =   4515
         TabIndex        =   3
         Top             =   525
         Width           =   1080
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Left            =   2730
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   540
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   300
         Left            =   1620
         TabIndex        =   1
         ToolTipText     =   "Informe a data quando ocorreu o relacionamento. Em caso de agendamento, informe a data de quando ocorrerá."
         Top             =   555
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data de Vigência:"
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
         Index           =   5
         Left            =   45
         TabIndex        =   31
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label DescricaoTabela 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3015
         TabIndex        =   30
         Top             =   180
         Width           =   5025
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tabela:"
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
         Index           =   9
         Left            =   855
         TabIndex        =   29
         Top             =   225
         Width           =   660
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8310
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   90
      Width           =   1140
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   60
         Picture         =   "TabelaPrecoGrupoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   570
         Picture         =   "TabelaPrecoGrupoOcx.ctx":0532
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4995
      Left            =   75
      TabIndex        =   32
      Top             =   960
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   8811
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Produtos\Preço"
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
Attribute VB_Name = "TabelaPrecoGrupoOcx"
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

Dim gobjTabelaPrecoGrupo As ClassTabelaPrecoGrupo

Dim objGridCategoria As AdmGrid
Dim iGrid_Categoria_Col As Integer
Dim iGrid_Valor_Col As Integer

Dim objGridProdutos As AdmGrid
Dim iGrid_ProdCodigo_Col As Integer
Dim iGrid_ProdDesc_Col As Integer
Dim iGrid_ProdUM_Col As Integer
Dim iGrid_ProdPrecoAntigo_Col As Integer
Dim iGrid_ProdPrecoNovo_Col As Integer

Private WithEvents objEventoProdutoPai As AdmEvento
Attribute objEventoProdutoPai.VB_VarHelpID = -1
Private WithEvents objEventoTipoDeProduto As AdmEvento
Attribute objEventoTipoDeProduto.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Tabela de Preço"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TabelaPrecoGrupo"

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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Activate()
    'Carrega os índices da tela
    'Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    'gi_ST_SetaIgnoraClick = 1
End Sub

Private Function Carrega_TabelaPreco() As Long
'Carrega a ComboBox Tabela

Dim lErro As Long
Dim iIndice As Integer
Dim colCodigo As New Collection

On Error GoTo Erro_Carrega_TabelaPreco

    'Preenche a ComboBox com  os Tipos de Documentos existentes no BD
    lErro = CF("TabelasPreco_Le_Codigos", colCodigo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    For iIndice = 1 To colCodigo.Count
        'Preenche a ComboBox Tabela com os objetos da colecao colTabelaPreco
        Tabela.AddItem colCodigo(iIndice)
        Tabela.ItemData(Tabela.NewIndex) = colCodigo(iIndice)
    Next

    Carrega_TabelaPreco = SUCESSO

    Exit Function

Erro_Carrega_TabelaPreco:

    Carrega_TabelaPreco = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208295)

    End Select

    Exit Function

End Function

Private Sub BotaoCriarTabela_Click()

Dim objTabelaPreco As New ClassTabelaPreco
Dim iIndice As Integer

    'Chama tela TabelaPrecoCriacao
    Call Chama_Tela_Modal("TabelaPrecoCriacao", objTabelaPreco)
    
    'Se não criou a Tabela -- > Sai
    If objTabelaPreco.iCodigo = 0 Then Exit Sub
        
    'Procura na Combo de Tabela o indice que a Tabela vai entrar
    If Tabela.ListCount > 0 Then
    
        For iIndice = 0 To Tabela.ListCount - 1
            If Tabela.ItemData(iIndice) > objTabelaPreco.iCodigo Then
                Exit For
            End If
        Next
    End If
    
    'Adiciona na Combo a Tabela Criada
    Tabela.AddItem objTabelaPreco.iCodigo, iIndice
    Tabela.ItemData(iIndice) = objTabelaPreco.iCodigo
    
    Tabela.ListIndex = iIndice
    
    Exit Sub

End Sub

Private Sub BotaoEditarTabela_Click()

Dim lErro As Long
Dim objTabelaPreco As New ClassTabelaPreco

On Error GoTo Error_BotaoEditarTabela_Click

    'Verifica se foi preenchida a ComboBox Tabela
    If Len(Trim(Tabela.Text)) = 0 Then gError 208296

    'Preenche os dados de objTabelaPreco a partir da tela
    objTabelaPreco.iCodigo = CInt(Tabela.Text)
    objTabelaPreco.sDescricao = DescricaoTabela.Caption

    'Chama tela TabelaPrecoAlteracao
    Call Chama_Tela_Modal("TabelaPrecoAlteracao", objTabelaPreco)

    'Coloca Descrição na tela
    DescricaoTabela.Caption = objTabelaPreco.sDescricao

    Exit Sub

Error_BotaoEditarTabela_Click:

    Select Case gErr

        Case 208296
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208297)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluirTabela_Click()

Dim lErro As Long
Dim objTabelaPreco As New ClassTabelaPreco
Dim vbMsgRes As VbMsgBoxResult
Dim colCodigo As New Collection
Dim iIndice As Integer

On Error GoTo Error_BotaoExcluirTabela_Click

    'Verifica se foi preenchida a ComboBox Tabela
    If Len(Trim(Tabela.Text)) = 0 Then gError 208298

    iIndice = Tabela.ListIndex
    
    objTabelaPreco.iCodigo = CInt(Tabela.Text)

    lErro = CF("TabelaPreco_Le", objTabelaPreco)
    If lErro <> SUCESSO And lErro <> 28004 Then gError ERRO_SEM_MENSAGEM

    If lErro = 28004 Then gError 208299

    'Pede confirmação da exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_TABELA_DE_PRECO", objTabelaPreco.iCodigo)

    'Se não confirmar, sai
    If vbMsgRes = vbNo Then Exit Sub

    'Exclui a Tabela de Preço
    lErro = CF("TabelaPreco_Exclui", objTabelaPreco)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Tabela.RemoveItem iIndice
    
    Tabela.ListIndex = -1
    DescricaoTabela.Caption = ""
    
    Exit Sub

Error_BotaoExcluirTabela_Click:

    Select Case gErr

        Case 208298
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)

        Case 208299
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_INEXISTENTE", gErr, objTabelaPreco.iCodigo)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208300)

    End Select

    Exit Sub

End Sub

Private Sub Data_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long
Dim lIntervalo As Long

On Error GoTo Erro_Data_Validate

    'Verifica se Data está preenchida
    If Len(Trim(Data.ClipText)) <> 0 Then

        'Critica a Data
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
               
    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208301)

    End Select

    Exit Sub

End Sub

Private Sub Data_GotFocus()
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)
End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208302)

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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208303)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoPai_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoPai_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoPai.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoPai.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoPai, "Gerencial = 1")

    Exit Sub

Erro_LabelProdutoPai_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208304)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoPai_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoPai_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoPai, DescProdPai)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Me.Show

    Exit Sub

Erro_objEventoProdutoPai_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208305)

    End Select

    Exit Sub

End Sub

Private Sub objEventoTipoDeProduto_evSelecao(obj1 As Object)

Dim objTipoProduto As ClassTipoDeProduto
Dim bCancel As Boolean

    Set objTipoProduto = obj1

    'coloca na tela o Tipo de Produto Selecionado e dispara o evento LostFocus
    TipoProduto.Text = objTipoProduto.iTipo
    Call TipoProduto_Validate(bCancel)

    Me.Show

End Sub

Public Sub LblTipoProduto_Click()

Dim objTipoDeProduto As ClassTipoDeProduto
Dim colSelecao As Collection

    Call Chama_Tela("TipoProdutoLista", colSelecao, objTipoDeProduto, objEventoTipoDeProduto)

End Sub

Public Sub TipoProduto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TipoProduto_GotFocus()
    Call MaskEdBox_TrataGotFocus(TipoProduto, iAlterado)
End Sub

Public Sub TipoProduto_Validate(Cancel As Boolean)
'Se mudar o tipo trazer dele os defaults para os campos da tela

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto

On Error GoTo Erro_TipoProduto_Validate

    If Len(Trim(TipoProduto.Text)) > 0 Then

        'Critica o valor
        lErro = Inteiro_Critica(TipoProduto.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        objTipoProduto.iTipo = CInt(TipoProduto.Text)
    
        'Lê o tipo
        lErro = CF("TipoDeProduto_Le", objTipoProduto)
        If lErro <> SUCESSO And lErro <> 22531 Then gError ERRO_SEM_MENSAGEM
        
        'Se não encontrar --> Erro
        If lErro = 22531 Then gError 208306
        
        DescTipoProduto.Caption = objTipoProduto.sDescricao
        
    Else
        DescTipoProduto.Caption = ""
    End If
    
    Exit Sub

Erro_TipoProduto_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 208306
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", gErr, TipoProduto.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208307)

    End Select

    Exit Sub

End Sub

Public Sub ProdutoPai_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ProdutoPai_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoPai_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoPai, DescProdPai)
    If lErro <> SUCESSO And lErro <> 27095 Then gError ERRO_SEM_MENSAGEM
    
    If lErro <> SUCESSO Then gError 208308

    Exit Sub

Erro_ProdutoPai_Validate:

    Cancel = True

    Select Case gErr

        Case 208308
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
       
        Case ERRO_SEM_MENSAGEM
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208309)

    End Select

    Exit Sub

End Sub

Public Sub Tabela_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Tabela_Click()

Dim lErro As Long
Dim objTabelaPreco As New ClassTabelaPreco

On Error GoTo Error_Tabela_Click

    'Verifica se foi preenchida a ComboBox Tabela
    If Tabela.ListIndex <> -1 Then

        objTabelaPreco.iCodigo = CInt(Tabela.Text)

        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError ERRO_SEM_MENSAGEM

        If lErro = 28004 Then gError 208310

        DescricaoTabela.Caption = objTabelaPreco.sDescricao

    End If

    iAlterado = 0

    Exit Sub

Error_Tabela_Click:

    Select Case gErr

        Case 208310
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_INEXISTENTE", gErr, objTabelaPreco.iCodigo)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208311)

    End Select

    Exit Sub

End Sub

Private Sub AnaliticoGrade_Click(Index As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Function Inicializa_Grid_Categoria(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Categoria

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Categoria")
    objGridInt.colColuna.Add ("Item")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ComboCategoriaProduto.Name)
    objGridInt.colCampo.Add (ComboCategoriaProdutoItem.Name)

    'Colunas do Grid
    iGrid_Categoria_Col = 1
    iGrid_Valor_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridCategoria

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 21

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridCategoria.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Categoria = SUCESSO

    Exit Function

End Function

Private Sub GridCategoria_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCategoria, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCategoria, iAlterado)
    End If

End Sub

Private Sub GridCategoria_GotFocus()

    Call Grid_Recebe_Foco(objGridCategoria)

End Sub

Private Sub GridCategoria_EnterCell()

    Call Grid_Entrada_Celula(objGridCategoria, iAlterado)

End Sub

Private Sub GridCategoria_LeaveCell()

    Call Saida_Celula(objGridCategoria)

End Sub

Private Sub GridCategoria_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridCategoria)

End Sub

Private Sub GridCategoria_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCategoria, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCategoria, iAlterado)
    End If

End Sub

Private Sub GridCategoria_LostFocus()

    Call Grid_Libera_Foco(objGridCategoria)

End Sub

Private Sub GridCategoria_RowColChange()

    Call Grid_RowColChange(objGridCategoria)

End Sub

Private Sub GridCategoria_Scroll()

    Call Grid_Scroll(objGridCategoria)

End Sub

Public Sub ComboCategoriaProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComboCategoriaProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

End Sub

Public Sub ComboCategoriaProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)

End Sub

Public Sub ComboCategoriaProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = ComboCategoriaProduto
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ComboCategoriaProdutoItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComboCategoriaProdutoItem_GotFocus()

Dim lErro As Long

On Error GoTo Erro_ComboCategoriaProdutoItem_GotFocus

    'Preenche com os ítens relacionados a Categoria correspondente
    Call Trata_ComboCategoriaProdutoItem

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

    Exit Sub

Erro_ComboCategoriaProdutoItem_GotFocus:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208312)

    End Select

    Exit Sub

End Sub

Public Sub ComboCategoriaProdutoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)

End Sub

Public Sub ComboCategoriaProdutoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = ComboCategoriaProdutoItem
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Trata_ComboCategoriaProdutoItem()

Dim lErro As Long
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim iIndice As Integer
Dim sValor As String

On Error GoTo Erro_Trata_ComboCategoriaProdutoItem

    sValor = ComboCategoriaProdutoItem.Text

    ComboCategoriaProdutoItem.Clear

    ComboCategoriaProdutoItem.Text = sValor

    'Se alguém estiver selecionado
    If Len(GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)) > 0 Then

        'Preencher a Combo de Itens desta Categoria
        objCategoriaProduto.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)

        lErro = Carrega_ComboCategoriaProdutoItem(objCategoriaProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        For iIndice = 0 To ComboCategoriaProdutoItem.ListCount - 1
            If ComboCategoriaProdutoItem.List(iIndice) = GridCategoria.Text Then
                ComboCategoriaProdutoItem.ListIndex = iIndice
                Exit For
            End If
        Next

    End If

    Exit Sub

Erro_Trata_ComboCategoriaProdutoItem:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208313)

    End Select
    
    Exit Sub

End Sub

Private Function Carrega_ComboCategoriaProdutoItem(objCategoriaProduto As ClassCategoriaProduto) As Long
'Carrega o Item da Categoria na Combobox

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem

On Error GoTo Erro_Carrega_ComboCategoriaProdutoItem

    'Lê a tabela CategoriaProdutoItem a partir da Categoria
    lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colItensCategoria)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Insere na combo CategoriaProdutoItem
    For Each objCategoriaProdutoItem In colItensCategoria
        'Insere na combo CategoriaProduto
        ComboCategoriaProdutoItem.AddItem objCategoriaProdutoItem.sItem
    Next

    Carrega_ComboCategoriaProdutoItem = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaProdutoItem:

    Carrega_ComboCategoriaProdutoItem = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208314)

    End Select

    Exit Function

End Function

Private Function Carrega_ComboCategoriaProduto() As Long
'Carrega as Categorias na Combobox

Dim lErro As Long
Dim colCategorias As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Carrega_ComboCategoriaProduto

    'Lê o código e a descrição de todas as categorias
    lErro = CF("CategoriasProduto_Le_Todas", colCategorias)
    If lErro <> SUCESSO And lErro <> 22542 Then gError ERRO_SEM_MENSAGEM

    For Each objCategoriaProduto In colCategorias
        'Insere na combo CategoriaProduto
        ComboCategoriaProduto.AddItem objCategoriaProduto.sCategoria
    Next

    Carrega_ComboCategoriaProduto = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaProduto:

    Carrega_ComboCategoriaProduto = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208315)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        If objGridInt.objGrid.Name = GridCategoria.Name Then

            'Verifica qual a coluna do Grid
            Select Case GridCategoria.Col

                Case iGrid_Categoria_Col

                    lErro = Saida_Celula_Categoria(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                Case iGrid_Valor_Col

                    lErro = Saida_Celula_Valor(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            End Select

        ElseIf objGridInt.objGrid.Name = GridProdutos.Name Then
        
            'Verifica qual a coluna do Grid
            Select Case GridProdutos.Col

                Case iGrid_ProdPrecoNovo_Col

                    lErro = Saida_Celula_ProdPrecoNovo(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            End Select
            
        End If


        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 208316

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 208316
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208317)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Categoria(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Saida_Celula_Categoria

    Set objGridInt.objControle = ComboCategoriaProduto

    If Len(Trim(ComboCategoriaProduto.Text)) > 0 Then

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ComboCategoriaProduto)
        If lErro <> SUCESSO Then
        
            'Preenche o objeto com a Categoria
             objCategoriaProduto.sCategoria = ComboCategoriaProduto.Text

             'Lê Categoria De Produto no BD
             lErro = CF("CategoriaProduto_Le", objCategoriaProduto)
             If lErro <> SUCESSO And lErro <> 22540 Then gError ERRO_SEM_MENSAGEM

             If lErro <> SUCESSO Then gError 208318  'Categoria não está cadastrada

        End If

        'Verifica se já existe a categoria no Grid
        For iIndice = 1 To objGridCategoria.iLinhasExistentes

            If iIndice <> GridCategoria.Row Then If GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col) = ComboCategoriaProduto.Text Then gError 208319

        Next

        If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    Else
        
        GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Valor_Col) = ""
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 208319
    
    Call Trata_TextoGrade

    Saida_Celula_Categoria = SUCESSO

    Exit Function

Erro_Saida_Celula_Categoria:

    Saida_Celula_Categoria = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 208318  'Categoria não está cadastrada

            'pergunta se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CATEGORIAPRODUTO")

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a Tela "CategoriaProduto"
                Call Chama_Tela("CategoriaProduto", objCategoriaProduto)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 208319
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_JA_SELECIONADA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208320)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Item do grid que está deixando de ser a corrente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colItens As New Collection

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridCategoria.objControle = ComboCategoriaProdutoItem

    If Len(Trim(ComboCategoriaProdutoItem.Text)) > 0 Then

        'se o campo de categoria estiver vazio ==> erro
        If Len(GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)) = 0 Then gError 208321

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ComboCategoriaProdutoItem)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
            objCategoriaProdutoItem.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)
            objCategoriaProdutoItem.sItem = ComboCategoriaProdutoItem.Text

            'Lê Categoria De Produto no BD
            lErro = CF("CategoriaProduto_Le_Item", objCategoriaProdutoItem)
            If lErro <> SUCESSO And lErro <> 22603 Then gError ERRO_SEM_MENSAGEM

            If lErro <> SUCESSO Then gError 208322 'Item da Categoria não está cadastrado

        End If

        If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Trata_TextoGrade

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 208321
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_CATEGORIA_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        

        Case 208322 'Item da Categoria não está cadastrado
            'Se não for perguntar se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CATEGORIAPRODUTOITEM")

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Preenche o objeto com a Categoria
                objCategoriaProduto.sCategoria = ComboCategoriaProduto.Text

                'Chama a Tela "CategoriaProduto"
                Call Chama_Tela("CategoriaProduto", objCategoriaProduto)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208323)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub CodigoLike_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DescricaoLike_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ModeloLike_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NomeRedLike_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ReferenciaLike_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TextoGrade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PrecoNovoPerc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PrecoNovoPerc_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PrecoNovoPerc_Validate

    'Veifica se CargaMax está preenchida
    If Len(Trim(PrecoNovoPerc.Text)) <> 0 Then

       'Critica a CargaMax
       lErro = Porcentagem_Critica(PrecoNovoPerc.Text)
       If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_PrecoNovoPerc_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208324)

    End Select

    Exit Sub

End Sub

Private Sub PrecoNovoRS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PrecoNovoRS_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PrecoNovoRS_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(PrecoNovoRS.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_Positivo_Critica(PrecoNovoRS.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    PrecoNovoRS.Text = Format(PrecoNovoRS.Text, PrecoNovoRS.Format)

    Exit Sub

Erro_PrecoNovoRS_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208325)

    End Select

    Exit Sub

End Sub

Public Sub Opcao_Click()

Dim lErro As Long
Dim objTabelaPrecoGrupo As New ClassTabelaPrecoGrupo

On Error GoTo Erro_Opcao_Click

    If Opcao.SelectedItem.Index <> iFrameAtual Then
        
        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
        
        'Se o frame anterior foi o de Seleção e ele foi alterado
        If iFrameAtual <> 1 Then
    
            DoEvents
            
            Call Grid_Limpa(objGridProdutos)
    
            lErro = Move_Selecao_Memoria(objTabelaPrecoGrupo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = CF("TabelaPrecoGrupo_Le", objTabelaPrecoGrupo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
              
            If objTabelaPrecoGrupo.colItens.Count = 0 Then gError 208326
    
            lErro = Traz_Produtos_Tela(objTabelaPrecoGrupo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            Set gobjTabelaPrecoGrupo = objTabelaPrecoGrupo
    
        End If
    
    End If

    Exit Sub

Erro_Opcao_Click:

    Select Case gErr
    
        Case 208326
            Call Rotina_Erro(vbOKOnly, "ERRO_SELECAO_NENHUM_PRODUTO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208327)

    End Select

    Exit Sub

End Sub

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Private Function Inicializa_Grid_Produtos(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Categoria

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Código")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("UM")
    objGridInt.colColuna.Add ("Preço Atual")
    objGridInt.colColuna.Add ("Preço Novo")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ProdCodigo.Name)
    objGridInt.colCampo.Add (ProdDesc.Name)
    objGridInt.colCampo.Add (ProdUM.Name)
    objGridInt.colCampo.Add (ProdPrecoAntigo.Name)
    objGridInt.colCampo.Add (ProdPrecoNovo.Name)

    'Colunas do Grid
    iGrid_ProdCodigo_Col = 1
    iGrid_ProdDesc_Col = 2
    iGrid_ProdUM_Col = 3
    iGrid_ProdPrecoAntigo_Col = 4
    iGrid_ProdPrecoNovo_Col = 5

    'Grid do GridInterno
    objGridInt.objGrid = GridProdutos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 1001

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 12

    'Largura da primeira coluna
    GridProdutos.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Produtos = SUCESSO

    Exit Function

End Function

Public Sub GridProdutos_Click()

Dim iExecutaEntradaCelula As Integer
Dim colcolColecao As New Collection

    Call Grid_Click(objGridProdutos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdutos, iAlterado)
    End If
    
    colcolColecao.Add gobjTabelaPrecoGrupo.colItens
    
    Call Ordenacao_ClickGrid(objGridProdutos, , colcolColecao)


End Sub

Public Sub GridProdutos_GotFocus()
    Call Grid_Recebe_Foco(objGridProdutos)
End Sub

Public Sub GridProdutos_EnterCell()
    Call Grid_Entrada_Celula(objGridProdutos, iAlterado)
End Sub

Public Sub GridProdutos_LeaveCell()
    Call Saida_Celula(objGridProdutos)
End Sub

Public Sub GridProdutos_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridProdutos)
End Sub

Public Sub GridProdutos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridProdutos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdutos, iAlterado)
    End If

End Sub

Public Sub GridProdutos_LostFocus()
    Call Grid_Libera_Foco(objGridProdutos)
End Sub

Public Sub GridProdutos_RowColChange()

Dim objTabelaPrecoGrupoItem As ClassTabelaPrecoGrupoItem
    
    Call Grid_RowColChange(objGridProdutos)
    
    Descricao.Caption = ""
    If GridProdutos.Row <> 0 And GridProdutos.Row <= objGridProdutos.iLinhasExistentes Then
        If Not (gobjTabelaPrecoGrupo Is Nothing) Then
            If gobjTabelaPrecoGrupo.colItens.Count >= GridProdutos.Row Then
                Set objTabelaPrecoGrupoItem = gobjTabelaPrecoGrupo.colItens.Item(GridProdutos.Row)
                Descricao.Caption = objTabelaPrecoGrupoItem.sDescricao
            End If
        End If
    End If

End Sub

Public Sub GridProdutos_Scroll()
    Call Grid_Scroll(objGridProdutos)
End Sub

Public Sub ProdPrecoNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ProdPrecoNovo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridProdutos)
End Sub

Public Sub ProdPrecoNovo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)
End Sub

Public Sub ProdPrecoNovo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = ProdPrecoNovo
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoProdutoPai = Nothing
    Set objEventoTipoDeProduto = Nothing
    Set objGridProdutos = Nothing
    Set objGridCategoria = Nothing
    Set gobjTabelaPrecoGrupo = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208328)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set gobjTabelaPrecoGrupo = New ClassTabelaPrecoGrupo
    Set objGridCategoria = New AdmGrid
    Set objGridProdutos = New AdmGrid
    Set objEventoProdutoPai = New AdmEvento
    Set objEventoTipoDeProduto = New AdmEvento

    iFrameAtual = 1

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoPai)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Chama Carrega_TabelaPreco
    lErro = Carrega_TabelaPreco()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Inicializa_Grid_Categoria(objGridCategoria)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Inicializa_Grid_Produtos(objGridProdutos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Carrega a combobox de Categoria Produto
    lErro = Carrega_ComboCategoriaProduto()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Data.PromptInclude = False
    Data.Text = Format(Date, "dd/mm/yy")
    Data.PromptInclude = True
    
    PrecoNovoRS.Format = gobjFAT.sFormatoPrecoUnitario
    
    If gobjCRFAT.iSeparaItensGradePrecoDif = DESMARCADO Then
        AnaliticoGrade(2).Value = vbUnchecked
        AnaliticoGrade(2).Enabled = False
        TextoGrade.Enabled = False
    Else
        AnaliticoGrade(2).Value = vbChecked
        AnaliticoGrade(2).Enabled = True
        TextoGrade.Enabled = True
    End If

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208329)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros() As Long
'
End Function

Function Move_Tela_Memoria() As Long

Dim lErro As Long, iIndice As Integer
Dim objTabelaPrecoGrupoItem As ClassTabelaPrecoGrupoItem

On Error GoTo Erro_Move_Tela_Memoria

    iIndice = 0
    For Each objTabelaPrecoGrupoItem In gobjTabelaPrecoGrupo.colItens
        iIndice = iIndice + 1
        objTabelaPrecoGrupoItem.dPrecoNovo = StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_ProdPrecoNovo_Col))
        objTabelaPrecoGrupoItem.sTextoGrade = TextoGrade.Text
    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208330)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long
'
End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long
'
End Function

Function Gravar_Registro() As Long

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If gobjTabelaPrecoGrupo.colItens.Count = 0 Then gError 208338

    lErro = Move_Tela_Memoria()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("TabelaPrecoGrupo_Grava", gobjTabelaPrecoGrupo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 208338
            Call Rotina_Erro(vbOKOnly, "ERRO_SELECAO_NENHUM_PRODUTO", gErr)
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208331)

    End Select

    Exit Function

End Function

Function Limpa_Tela_TabelaPreco() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TabelaPreco

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    Call Limpa_Tela_TabelaPreco2
    
    Call Grid_Limpa(objGridCategoria)
    
    Tabela.ListIndex = -1
    
    DescricaoTabela.Caption = ""
    Descricao.Caption = ""
    DescTipoProduto.Caption = ""
    DescProdPai.Caption = ""
    
    AnaliticoGrade(0).Value = vbChecked
    AnaliticoGrade(1).Value = vbChecked
    
    If gobjCRFAT.iSeparaItensGradePrecoDif = DESMARCADO Then
        AnaliticoGrade(2).Value = vbUnchecked
    Else
        AnaliticoGrade(2).Value = vbChecked
    End If

    OptPreco(0).Value = True
    
    Data.PromptInclude = False
    Data.Text = Format(Date, "dd/mm/yy")
    Data.PromptInclude = True
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0

    Limpa_Tela_TabelaPreco = SUCESSO

    Exit Function

Erro_Limpa_Tela_TabelaPreco:

    Limpa_Tela_TabelaPreco = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208332)

    End Select

    Exit Function

End Function

Function Limpa_Tela_TabelaPreco2() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TabelaPreco2

    Call Grid_Limpa(objGridProdutos)
    
    Call Ordenacao_Limpa(objGridProdutos)
    
    'Torna Frame atual invisível
    Frame1(Opcao.SelectedItem.Index).Visible = False
    iFrameAtual = 1
    'Torna Frame atual visível
    Frame1(iFrameAtual).Visible = True
    Opcao.Tabs.Item(iFrameAtual).Selected = True
    
    Call Opcao_Click

    Limpa_Tela_TabelaPreco2 = SUCESSO

    Exit Function

Erro_Limpa_Tela_TabelaPreco2:

    Limpa_Tela_TabelaPreco2 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208333)

    End Select

    Exit Function

End Function

Function Traz_Produtos_Tela(ByVal objTabelaPrecoGrupo As ClassTabelaPrecoGrupo) As Long

Dim lErro As Long, iIndice As Integer, sProdMask As String
Dim objTabelaPrecoGrupoItem As ClassTabelaPrecoGrupoItem

On Error GoTo Erro_Traz_Produtos_Tela
    
    Call Grid_Limpa(objGridProdutos)
    
    Call Ordenacao_Limpa(objGridProdutos)
    
    iIndice = 0
    For Each objTabelaPrecoGrupoItem In objTabelaPrecoGrupo.colItens
        iIndice = iIndice + 1
        
        Call Mascara_RetornaProdutoTela(objTabelaPrecoGrupoItem.sProduto, sProdMask)
        
        GridProdutos.TextMatrix(iIndice, iGrid_ProdCodigo_Col) = sProdMask
        GridProdutos.TextMatrix(iIndice, iGrid_ProdDesc_Col) = objTabelaPrecoGrupoItem.sDescricao
        GridProdutos.TextMatrix(iIndice, iGrid_ProdUM_Col) = objTabelaPrecoGrupoItem.sUM
        If objTabelaPrecoGrupoItem.dPrecoAtual <> 0 Then GridProdutos.TextMatrix(iIndice, iGrid_ProdPrecoAntigo_Col) = Format(objTabelaPrecoGrupoItem.dPrecoAtual, gobjFAT.sFormatoPrecoUnitario)
        If objTabelaPrecoGrupoItem.dPrecoNovo <> 0 Then GridProdutos.TextMatrix(iIndice, iGrid_ProdPrecoNovo_Col) = Format(objTabelaPrecoGrupoItem.dPrecoNovo, gobjFAT.sFormatoPrecoUnitario)

    Next
    
    objGridProdutos.iLinhasExistentes = objTabelaPrecoGrupo.colItens.Count
    
    iAlterado = 0

    Traz_Produtos_Tela = SUCESSO

    Exit Function

Erro_Traz_Produtos_Tela:

    Traz_Produtos_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208334)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Limpa Tela
    Call Limpa_Tela_TabelaPreco2

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208335)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208336)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Limpa_Tela_TabelaPreco

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208337)

    End Select

    Exit Sub

End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
       
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is TipoProduto Then
            Call LblTipoProduto_Click
        ElseIf Me.ActiveControl Is ProdutoPai Then
            Call LabelProdutoPai_Click
        End If
    
    End If

End Sub

Function Move_Selecao_Memoria(ByVal objTabelaPrecoGrupo As ClassTabelaPrecoGrupo) As Long
'Recolhe dados do TAB de Seleção

Dim lErro As Long
Dim iIndice As Integer, sProduto As String, iPreenchido As Integer
Dim objProdCat As ClassProdutoCategoria

On Error GoTo Erro_Move_Selecao_Memoria

    objTabelaPrecoGrupo.iTabela = StrParaInt(Tabela.Text)
    objTabelaPrecoGrupo.dtDataVigencia = StrParaDate(Data.Text)
    
    If objTabelaPrecoGrupo.iTabela = 0 Then gError 208339
    If objTabelaPrecoGrupo.dtDataVigencia = DATA_NULA Then gError 208340
    If objTabelaPrecoGrupo.dtDataVigencia < Date Then gError 208341

    objTabelaPrecoGrupo.dtDataRef = objTabelaPrecoGrupo.dtDataVigencia
    
    objTabelaPrecoGrupo.dPrecoNovoPerc = StrParaDbl(Val(PrecoNovoPerc.Text) / 100)
    objTabelaPrecoGrupo.dPrecoNovoRS = StrParaDbl(PrecoNovoRS.Text)
    objTabelaPrecoGrupo.iFilialEmpresa = giFilialEmpresa
    objTabelaPrecoGrupo.iTipoDeProduto = StrParaInt(TipoProduto.Text)
    
    If AnaliticoGrade(0) = vbChecked Then
        objTabelaPrecoGrupo.iAnaliticoSemGrade = MARCADO
    Else
        objTabelaPrecoGrupo.iAnaliticoSemGrade = DESMARCADO
    End If
    
    If AnaliticoGrade(1) = vbChecked Then
        objTabelaPrecoGrupo.iGradeKitVenda = MARCADO
    Else
        objTabelaPrecoGrupo.iGradeKitVenda = DESMARCADO
    End If
    
    If AnaliticoGrade(2) = vbChecked Then
        objTabelaPrecoGrupo.iAnaliticoComGrade = MARCADO
    Else
        objTabelaPrecoGrupo.iAnaliticoComGrade = DESMARCADO
    End If
    
    If OptPreco(PRECO_GRUPO_TIPO_VALOR) Then
        objTabelaPrecoGrupo.iTipoNovoPreco = PRECO_GRUPO_TIPO_VALOR
    Else
        objTabelaPrecoGrupo.iTipoNovoPreco = PRECO_GRUPO_TIPO_PERCENTUAL
    End If
    
    objTabelaPrecoGrupo.sCodigoLike = CodigoLike.Text
    objTabelaPrecoGrupo.sDescricaoLike = DescricaoLike.Text
    objTabelaPrecoGrupo.sModeloLike = ModeloLike.Text
    objTabelaPrecoGrupo.sNomeRedLike = NomeRedLike.Text
    objTabelaPrecoGrupo.sReferenciaLike = ReferenciaLike.Text
   
    If Len(Trim(ProdutoPai.ClipText)) <> 0 Then
   
        lErro = CF("Produto_Formata", ProdutoPai.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
   
        objTabelaPrecoGrupo.sProdutoPai = sProduto
        
    End If
    
    For iIndice = 1 To objGridCategoria.iLinhasExistentes
        Set objProdCat = New ClassProdutoCategoria
        
        objProdCat.sCategoria = GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col)
        objProdCat.sItem = GridCategoria.TextMatrix(iIndice, iGrid_Valor_Col)
    
        objTabelaPrecoGrupo.colCategorias.Add objProdCat
    Next
   
    Move_Selecao_Memoria = SUCESSO

    Exit Function

Erro_Move_Selecao_Memoria:

    Move_Selecao_Memoria = gErr

    Select Case gErr

        Case 208339
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)

        Case 208340
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_VIGENCIA_NAO_PREENCHIDA", gErr)

        Case 208341
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_VIGENCIA_MENOR_DATA_ATUAL", gErr, Date)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208342)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ProdPrecoNovo(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ProdPrecoNovo

    Set objGridInt.objControle = ProdPrecoNovo

    'Se estiver preenchido
    If Len(Trim(objGridInt.objControle.Text)) > 0 Then

        lErro = Valor_Positivo_Critica(objGridInt.objControle.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
   
    Saida_Celula_ProdPrecoNovo = SUCESSO

    Exit Function

Erro_Saida_Celula_ProdPrecoNovo:

    Saida_Celula_ProdPrecoNovo = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208343)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub OptPreco_Click(Index As Integer)
    iAlterado = REGISTRO_ALTERADO
    If OptPreco(PRECO_GRUPO_TIPO_VALOR) Then
        PrecoNovoPerc.Enabled = False
        PrecoNovoRS.Enabled = True
        PrecoNovoPerc.Text = ""
    Else
        PrecoNovoPerc.Enabled = True
        PrecoNovoRS.Enabled = False
        PrecoNovoRS.Text = ""
    End If
End Sub

Private Sub BotaoProduto_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_BotaoProduto_Click

    If GridProdutos.Row = 0 Then gError 208344
    
    objProduto.sCodigo = gobjTabelaPrecoGrupo.colItens.Item(GridProdutos.Row).sProduto

    Call Chama_Tela("Produto", objProduto)

    Exit Sub

Erro_BotaoProduto_Click:

    Select Case gErr
    
        Case 208344
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208345)

    End Select

    Exit Sub
    
End Sub

Private Sub Trata_TextoGrade()

Dim lErro As Long
Dim iIndice As Integer
Dim sTexto As String

On Error GoTo Erro_Trata_TextoGrade
    
    If gobjCRFAT.iSeparaItensGradePrecoDif = MARCADO Then
        
        sTexto = ""
        For iIndice = 1 To objGridCategoria.iLinhasExistentes
               
            sTexto = sTexto & IIf(Len(Trim(sTexto)) = 0, "", SEPARADOR) & GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col)
            sTexto = sTexto & " " & GridCategoria.TextMatrix(iIndice, iGrid_Valor_Col)
        
        Next
        
    End If
    
    TextoGrade.Text = sTexto
    
    Exit Sub

Erro_Trata_TextoGrade:

    Select Case gErr
    
          Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208345)

    End Select

    Exit Sub
    
End Sub
