VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpNFPedComprasOcx 
   ClientHeight    =   6570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   KeyPreview      =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   8055
   Begin VB.Frame Frame5 
      Caption         =   "Pedidos de Compra"
      Height          =   2076
      Left            =   96
      TabIndex        =   48
      Top             =   4284
      Width           =   7860
      Begin VB.Frame Frame9 
         Caption         =   "Categoria"
         Height          =   972
         Left            =   180
         TabIndex        =   55
         Top             =   960
         Width           =   7441
         Begin VB.Frame Frame10 
            Caption         =   "Item"
            Height          =   684
            Left            =   3240
            TabIndex        =   62
            Top             =   144
            Width           =   4035
            Begin VB.ComboBox ItemDe 
               Height          =   288
               Left            =   528
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   276
               Width           =   1428
            End
            Begin VB.ComboBox ItemAte 
               Height          =   288
               Left            =   2448
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   252
               Width           =   1428
            End
            Begin VB.Label Label10 
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
               Height          =   192
               Left            =   204
               TabIndex        =   64
               Top             =   312
               Width           =   312
            End
            Begin VB.Label Label9 
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
               Height          =   192
               Left            =   2076
               TabIndex        =   63
               Top             =   312
               Width           =   360
            End
         End
         Begin VB.ComboBox Categoria 
            Height          =   288
            Left            =   1044
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   348
            Width           =   2100
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Categoria:"
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
            Height          =   192
            Left            =   144
            TabIndex        =   61
            Top             =   390
            Width           =   876
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Produtos"
         Height          =   672
         Left            =   4332
         TabIndex        =   52
         Top             =   225
         Width           =   3270
         Begin MSMask.MaskEdBox ProdutoDe 
            Height          =   300
            Left            =   516
            TabIndex        =   17
            Top             =   240
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoAte 
            Height          =   300
            Left            =   2112
            TabIndex        =   18
            Top             =   240
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin VB.Label LabelProdutoAte 
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
            Left            =   1695
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   54
            Top             =   300
            Width           =   360
         End
         Begin VB.Label LabelProdutoDe 
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
            TabIndex        =   53
            Top             =   315
            Width           =   315
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Comprador"
         Height          =   684
         Left            =   192
         TabIndex        =   49
         Top             =   218
         Width           =   4035
         Begin VB.ComboBox CompradorDe 
            Height          =   288
            Left            =   492
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   240
            Width           =   1464
         End
         Begin VB.ComboBox CompradorAte 
            Height          =   288
            Left            =   2412
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   216
            Width           =   1464
         End
         Begin VB.Label Label5 
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
            TabIndex        =   51
            Top             =   270
            Width           =   315
         End
         Begin VB.Label Label4 
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
            Left            =   2040
            TabIndex        =   50
            Top             =   270
            Width           =   360
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5820
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   108
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpNFPedComprasOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpNFPedComprasOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpNFPedComprasOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpNFPedComprasOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
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
      Left            =   4044
      Picture         =   "RelOpNFPedComprasOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   84
      Width           =   1635
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   288
      ItemData        =   "RelOpNFPedComprasOcx.ctx":0A96
      Left            =   792
      List            =   "RelOpNFPedComprasOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   3090
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filial Empresa"
      Height          =   732
      Left            =   96
      TabIndex        =   40
      Top             =   828
      Width           =   7860
      Begin VB.ComboBox FilialEmpresaAte 
         Height          =   288
         Left            =   4428
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   288
         Width           =   3180
      End
      Begin VB.ComboBox FilialEmpresaDe 
         Height          =   288
         Left            =   648
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   288
         Width           =   3180
      End
      Begin VB.Label LabelCodFilialDe 
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
         Height          =   192
         Left            =   288
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   47
         Top             =   324
         Width           =   312
      End
      Begin VB.Label LabelCodFilialAte 
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
         Height          =   192
         Left            =   4032
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   46
         Top             =   324
         Width           =   360
      End
   End
   Begin VB.ComboBox ComboOrdenacao 
      Height          =   315
      ItemData        =   "RelOpNFPedComprasOcx.ctx":0A9A
      Left            =   -20000
      List            =   "RelOpNFPedComprasOcx.ctx":0AA7
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   735
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Notas Fiscais"
      Height          =   2544
      Left            =   96
      TabIndex        =   28
      Top             =   1632
      Width           =   7860
      Begin VB.Frame Frame11 
         Caption         =   "Data de Atualização"
         Height          =   684
         Left            =   216
         TabIndex        =   56
         Top             =   1644
         Width           =   4035
         Begin MSComCtl2.UpDown UpDownAtualizacaoDe 
            Height          =   315
            Left            =   1665
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox AtualizacaoDe 
            Height          =   315
            Left            =   480
            TabIndex        =   12
            Top             =   255
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownAtualizacaoAte 
            Height          =   315
            Left            =   3630
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox AtualizacaoAte 
            Height          =   315
            Left            =   2445
            TabIndex        =   13
            Top             =   255
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label14 
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
            Left            =   2070
            TabIndex        =   60
            Top             =   315
            Width           =   360
         End
         Begin VB.Label Label13 
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
            TabIndex        =   59
            Top             =   315
            Width           =   315
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Séries"
         Height          =   645
         Left            =   216
         TabIndex        =   43
         Top             =   218
         Width           =   4035
         Begin VB.ComboBox SerieAte 
            Height          =   288
            Left            =   2448
            TabIndex        =   5
            Top             =   210
            Width           =   885
         End
         Begin VB.ComboBox SerieDe 
            Height          =   288
            Left            =   495
            TabIndex        =   4
            Top             =   240
            Width           =   885
         End
         Begin VB.Label LabelSerieAte 
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
            Height          =   192
            Left            =   2076
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   45
            Top             =   276
            Width           =   360
         End
         Begin VB.Label LabelSerieDe 
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
            TabIndex        =   44
            Top             =   270
            Width           =   315
         End
      End
      Begin VB.Frame FrameCodigo 
         Caption         =   "Número"
         Height          =   630
         Left            =   4368
         TabIndex        =   37
         Top             =   225
         Width           =   3270
         Begin MSMask.MaskEdBox NumeroDe 
            Height          =   300
            Left            =   510
            TabIndex        =   6
            Top             =   240
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumeroAte 
            Height          =   300
            Left            =   2115
            TabIndex        =   7
            Top             =   240
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin VB.Label LabelNumeroDe 
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
            TabIndex        =   39
            Top             =   315
            Width           =   315
         End
         Begin VB.Label LabelNumeroAte 
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
            Left            =   1695
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   38
            Top             =   300
            Width           =   360
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Data de Entrada"
         Height          =   684
         Left            =   216
         TabIndex        =   32
         Top             =   918
         Width           =   4035
         Begin MSComCtl2.UpDown UpDownDataDe 
            Height          =   315
            Left            =   1665
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataDe 
            Height          =   315
            Left            =   480
            TabIndex        =   8
            Top             =   255
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataAte 
            Height          =   315
            Left            =   3630
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataAte 
            Height          =   315
            Left            =   2445
            TabIndex        =   9
            Top             =   255
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label2 
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
            TabIndex        =   36
            Top             =   315
            Width           =   315
         End
         Begin VB.Label Label3 
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
            Left            =   2070
            TabIndex        =   35
            Top             =   315
            Width           =   360
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Fornecedores"
         Height          =   696
         Left            =   4368
         TabIndex        =   29
         Top             =   912
         Width           =   3270
         Begin MSMask.MaskEdBox FornecedorDe 
            Height          =   300
            Left            =   525
            TabIndex        =   10
            Top             =   255
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FornecedorAte 
            Height          =   300
            Left            =   2130
            TabIndex        =   11
            Top             =   255
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin VB.Label LabelFornecedorDe 
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
            TabIndex        =   31
            Top             =   315
            Width           =   315
         End
         Begin VB.Label LabelFornecedorAte 
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
            Left            =   1710
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   30
            Top             =   315
            Width           =   360
         End
      End
      Begin VB.CheckBox CheckItens 
         Caption         =   "Exibe Item a Item"
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
         Left            =   4755
         TabIndex        =   14
         Top             =   2040
         Width           =   1890
      End
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
      Height          =   252
      Left            =   144
      TabIndex        =   42
      Top             =   252
      Width           =   612
   End
   Begin VB.Label Label8 
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
      Left            =   -20000
      TabIndex        =   41
      Top             =   810
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "RelOpNFPedComprasOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***** Alteração feita em 27/04/01 ***********
'A combo Ordenacao não está visível em tempo de execução, pois o relatório não está preparado para ordenar
'Para tornar essa combo visível, será necessário verificar se o código para ordenação está correto,
'e alterar o relatório, deixando-o preparado para aceitar as possíveis ordenações
'***** Feito por Luiz Gustavo *****************


'??? ATENCAO: Quem for refazer o Relatório deverá prestar atenção as novas Macros S##
'Os valores foram trocados, pois alguns sairam e outros entraram ...
'Me pergunte ... (Daniel)


'Property Variables:
Dim m_Caption As String
Event Unload()

'RelOpCompradoresPC
Const ORD_POR_NF = 0
Const ORD_POR_DATA = 1
Const ORD_POR_FORNECEDOR = 2

Private WithEvents objEventoCodPCDe As AdmEvento
Attribute objEventoCodPCDe.VB_VarHelpID = -1
Private WithEvents objEventoCodPCAte As AdmEvento
Attribute objEventoCodPCAte.VB_VarHelpID = -1
Private WithEvents objEventoCodFilialDe As AdmEvento
Attribute objEventoCodFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoCodFilialAte As AdmEvento
Attribute objEventoCodFilialAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialDe As AdmEvento
Attribute objEventoNomeFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialAte As AdmEvento
Attribute objEventoNomeFilialAte.VB_VarHelpID = -1
Private WithEvents objEventoNumNFDe As AdmEvento
Attribute objEventoNumNFDe.VB_VarHelpID = -1
Private WithEvents objEventoNumNFAte As AdmEvento
Attribute objEventoNumNFAte.VB_VarHelpID = -1
Private WithEvents objEventoFornDe As AdmEvento
Attribute objEventoFornDe.VB_VarHelpID = -1
Private WithEvents objEventoFornAte As AdmEvento
Attribute objEventoFornAte.VB_VarHelpID = -1
Private WithEvents objEventoSerieDe As AdmEvento
Attribute objEventoSerieDe.VB_VarHelpID = -1
Private WithEvents objEventoSerieAte As AdmEvento
Attribute objEventoSerieAte.VB_VarHelpID = -1
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1


Dim iAlterado As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 73523

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 73524

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 73523

        Case 73524
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170239)

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
    If lErro <> SUCESSO Then gError 73525

    ComboOrdenacao.ListIndex = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    CheckItens.Value = vbUnchecked

    Exit Sub

Erro_Limpa_Tela_Rel:

    Select Case gErr

        Case 73525

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170240)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodPCDe = New AdmEvento
    Set objEventoCodPCAte = New AdmEvento
    Set objEventoCodFilialDe = New AdmEvento
    Set objEventoCodFilialAte = New AdmEvento
    Set objEventoNomeFilialDe = New AdmEvento
    Set objEventoNomeFilialAte = New AdmEvento
    Set objEventoFornDe = New AdmEvento
    Set objEventoFornAte = New AdmEvento
    Set objEventoSerieDe = New AdmEvento
    Set objEventoSerieAte = New AdmEvento
    Set objEventoNumNFDe = New AdmEvento
    Set objEventoNumNFAte = New AdmEvento
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    
    ComboOrdenacao.ListIndex = 0

    lErro = Carrega_Serie()
    If lErro <> SUCESSO Then gError 73526
    
    lErro = Carrega_FilialEmpresa()
    If lErro <> SUCESSO Then gError 108858
    
    lErro = Carrega_Compradores()
    If lErro <> SUCESSO Then gError 108859
    
    lErro = Carrega_Categorias()
    If lErro <> SUCESSO Then gError 108860
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoDe)
    If lErro <> SUCESSO Then gError 108861

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoAte)
    If lErro <> SUCESSO Then gError 108862
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 73526, 108858 To 108862
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170241)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

    Set objEventoCodPCDe = Nothing
    Set objEventoCodPCAte = Nothing
    Set objEventoCodFilialDe = Nothing
    Set objEventoCodFilialAte = Nothing
    Set objEventoNomeFilialDe = Nothing
    Set objEventoNomeFilialAte = Nothing
    Set objEventoNumNFDe = Nothing
    Set objEventoNumNFAte = Nothing
    Set objEventoFornDe = Nothing
    Set objEventoFornAte = Nothing
    Set objEventoSerieDe = Nothing
    Set objEventoSerieAte = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    
End Sub

Private Sub Categoria_Click()
'Preenche os itens da categoria selecionada

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem

On Error GoTo Erro_Categoria_Click

    'Preenche o Obj
    objCategoriaProduto.sCategoria = Categoria.List(Categoria.ListIndex)
    
    'Le as categorias do Produto
    lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colItensCategoria)
    If lErro <> SUCESSO And lErro <> 22541 Then gError 108885
    
    ItemDe.Clear
    ItemAte.Clear
    
    ItemDe.AddItem ("")
    ItemAte.AddItem ("")
    
    For Each objCategoriaProdutoItem In colItensCategoria
        
        ItemDe.AddItem (objCategoriaProdutoItem.sItem)
        ItemAte.AddItem (objCategoriaProdutoItem.sItem)
        
    Next

    Exit Sub

Erro_Categoria_Click:

    Select Case gErr

         Case 108885
         
         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170242)

    End Select

End Sub

Private Sub DataAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)
    
End Sub

Private Sub DataDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)
    
End Sub

Private Sub FornecedorAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(FornecedorAte, iAlterado)
    
End Sub

Private Sub FornecedorDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(FornecedorDe, iAlterado)
    
End Sub

Private Sub LabelSerieDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objSerie As New ClassSerie

On Error GoTo Erro_LabelSerieDe_Click

    If Len(Trim(SerieDe.Text)) > 0 Then
    
        objSerie.sSerie = SerieDe.Text
    End If
    
    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerieDe)
    
    Exit Sub

Erro_LabelSerieDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170243)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelSerieAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objSerie As New ClassSerie

On Error GoTo Erro_LabelSerieAte_Click

    If Len(Trim(SerieAte.Text)) > 0 Then
    
        objSerie.sSerie = SerieAte.Text
    End If
    
    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerieAte)
    
    Exit Sub

Erro_LabelSerieAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170244)

    End Select

    Exit Sub
    
End Sub


Private Sub LabelFornecedorDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelFornecedorDe_Click

    If Len(Trim(FornecedorDe.Text)) > 0 Then
    
        objFornecedor.lCodigo = StrParaLong(FornecedorDe.Text)
        
    End If
    
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornDe)
    
    Exit Sub

Erro_LabelFornecedorDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170245)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelFornecedorAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelFornecedorAte_Click

    If Len(Trim(FornecedorAte.Text)) > 0 Then
    
        objFornecedor.lCodigo = StrParaLong(FornecedorAte.Text)
        
    End If
    
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornAte)
    
    Exit Sub

Erro_LabelFornecedorAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170246)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelNumeroAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objNF As New ClassNFiscal

On Error GoTo Erro_LabelNumeroAte_Click

    If Len(Trim(NumeroAte.Text)) > 0 Then
        'Preenche com o numero da tela
        objNF.lNumNotaFiscal = StrParaLong(NumeroAte.Text)
    End If

    'Chama Tela NFiscalEntradaTodasLista
    Call Chama_Tela("NFiscalEntradaTodasLista", colSelecao, objNF, objEventoNumNFAte)

   Exit Sub

Erro_LabelNumeroAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170247)

    End Select

    Exit Sub

End Sub

Private Sub LabelNumeroDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objNF As New ClassNFiscal

On Error GoTo Erro_LabelNumeroDe_Click

    If Len(Trim(NumeroDe.Text)) > 0 Then
        'Preenche com o numero da tela
        objNF.lNumNotaFiscal = StrParaLong(NumeroDe.Text)
    End If

    'Chama Tela NFiscalEntradaTodasLista
    Call Chama_Tela("NFiscalEntradaTodasLista", colSelecao, objNF, objEventoNumNFDe)

   Exit Sub

Erro_LabelNumeroDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170248)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataDe.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 73527

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 73527
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170249)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataAte.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 73528

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 73528
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170250)

    End Select

    Exit Sub

End Sub

Private Sub NumeroAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(NumeroAte, iAlterado)
    
End Sub

Private Sub NumeroDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(NumeroDe, iAlterado)
    
End Sub

Private Sub SerieDe_Validate(Cancel As Boolean)

Dim objSerie As New ClassSerie
Dim lErro As Long

On Error GoTo Erro_SerieDe_Validate

    If Len(Trim(SerieDe.Text)) > 0 Then
        
        objSerie.sSerie = SerieDe.Text
        
        lErro = CF("Serie_Le", objSerie)
        If lErro <> SUCESSO Then gError 73681
        
    End If
    
    Exit Sub
    
Erro_SerieDe_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 73681
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170251)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub SerieAte_Validate(Cancel As Boolean)

Dim objSerie As New ClassSerie
Dim lErro As Long

On Error GoTo Erro_SerieAte_Validate

    If Len(Trim(SerieAte.Text)) > 0 Then
        
        objSerie.sSerie = SerieAte.Text
        
        lErro = CF("Serie_Le", objSerie)
        If lErro <> SUCESSO Then gError 73682
        
    End If
    
    Exit Sub
    
Erro_SerieAte_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 73682
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170252)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 73529

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 73529
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170253)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 73530

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 73530
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170254)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 73531

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 73531
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170255)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 73532

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 73532
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170256)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFornDe_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    FornecedorDe.Text = CStr(objFornecedor.lCodigo)

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoFornAte_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    FornecedorAte.Text = CStr(objFornecedor.lCodigo)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoSerieAte_evSelecao(obj1 As Object)

Dim objSerie As New ClassSerie

    Set objSerie = obj1

    SerieAte.Text = objSerie.sSerie

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoSerieDe_evSelecao(obj1 As Object)

Dim objSerie As New ClassSerie

    Set objSerie = obj1

    SerieDe.Text = objSerie.sSerie

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNumNFAte_evSelecao(obj1 As Object)

Dim objNF As New ClassNFiscal

    Set objNF = obj1

    NumeroAte.Text = CStr(objNF.lNumNotaFiscal)

    Me.Show

End Sub

Private Sub objEventoNumNFDe_evSelecao(obj1 As Object)

Dim objNF As New ClassNFiscal

    Set objNF = obj1

    NumeroDe.Text = CStr(objNF.lNumNotaFiscal)

    Me.Show

End Sub


Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 73534

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 73535

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 73536

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 73537

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 73534
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 73535 To 73537

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170257)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 73538

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPNFPEDCOM")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 73539

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 73538
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 73539

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170258)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 73540
    
'    Select Case ComboOrdenacao.ListIndex
'
'            Case ORD_POR_NF
'
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "NFSerie", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "NFNumero", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataEntrada", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "PedCompraCod", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemPedCompra", 1)
'
'            Case ORD_POR_DATA
'
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataEntrada", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "NFSerie", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "NFNumero", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "PedCompraCod", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemPedCompra", 1)
'
'            Case ORD_POR_FORNECEDOR
'
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "FornecedorCod", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilFornCod", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "NFSerie", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "NFNumero", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataEntrada", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "PedCompraCod", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemPedCompra", 1)
'
'            Case Else
'                gError 74949
'
'    End Select

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 73540, 74949

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170259)

    End Select

    Exit Sub

End Sub


Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim iFilial_I As Integer, iFilial_F As Integer, sNomeFil_I As String, sNomeFil_F As String, sNumero_I As String
Dim sNumero_F As String, sSerie_I As String, sSerie_F As String, sForn_I As String, sForn_F As String, sCheck As String
Dim sOrdenacaoPor As String, iOrdenacao As Long, sOrd As String
Dim sComprador_I As String, sComprador_F As String, dtAtualizacao_I As Date, dtAtualizacao_F As Date
Dim sProduto_I As String, sProduto_F As String, sCategoria As String, sItem_I As String, sItem_F As String

On Error GoTo Erro_PreencherRelOp

    lErro = Formata_E_Critica_Parametros(iFilial_I, iFilial_F, sNomeFil_I, sNomeFil_F, sSerie_I, sSerie_F, sNumero_I, sNumero_F, sForn_I, sForn_F, sComprador_I, sComprador_F, sProduto_I, sProduto_F, sItem_I, sItem_F)
    If lErro <> SUCESSO Then gError 73541

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 73542

    lErro = objRelOpcoes.IncluirParametro("NCODFILIALINIC", CStr(iFilial_I))
    If lErro <> AD_BOOL_TRUE Then gError 73543

    lErro = objRelOpcoes.IncluirParametro("NCODFILIALFIM", CStr(iFilial_F))
    If lErro <> AD_BOOL_TRUE Then gError 73549
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALINIC", sNomeFil_I)
    If lErro <> AD_BOOL_TRUE Then gError 73544

    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALFIM", sNomeFil_F)
    If lErro <> AD_BOOL_TRUE Then gError 73550
    
    lErro = objRelOpcoes.IncluirParametro("NNOTAFISCALINIC", sNumero_I)
    If lErro <> AD_BOOL_TRUE Then gError 73545

    lErro = objRelOpcoes.IncluirParametro("NCODFORNINIC", sForn_I)
    If lErro <> AD_BOOL_TRUE Then gError 73546

    lErro = objRelOpcoes.IncluirParametro("TSERIEINIC", sSerie_I)
    If lErro <> AD_BOOL_TRUE Then gError 73547

    'Preenche data inicial
    If Trim(DataDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAINIC", DataDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 73548

    lErro = objRelOpcoes.IncluirParametro("NNOTAFISCALFIM", sNumero_F)
    If lErro <> AD_BOOL_TRUE Then gError 73551

    lErro = objRelOpcoes.IncluirParametro("NCODFORNFIM", sForn_F)
    If lErro <> AD_BOOL_TRUE Then gError 73552

    lErro = objRelOpcoes.IncluirParametro("TSERIEFIM", sSerie_F)
    If lErro <> AD_BOOL_TRUE Then gError 73553

    'Preenche data final
    If Trim(DataAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAFIM", DataAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 73554

    'Exibe Itens
    If CheckItens.Value = 0 Then
        sCheck = 0
        gobjRelatorio.sNomeTsk = "nfxpc"
    Else
        sCheck = 1
        gobjRelatorio.sNomeTsk = "nfxpcit"
    End If

    lErro = objRelOpcoes.IncluirParametro("NITENS", sCheck)
    If lErro <> AD_BOOL_TRUE Then gError 73555
    
    'Preenche data atualizacao
    If Trim(AtualizacaoDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DTATUINI", CDate(AtualizacaoDe.Text))
    Else
        lErro = objRelOpcoes.IncluirParametro("DTATUINI", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 108900
    
    'Preenche data atualizacao
    If Trim(AtualizacaoAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DTATUFIM", CDate(AtualizacaoAte.Text))
    Else
        lErro = objRelOpcoes.IncluirParametro("DTATUFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 108901
    
    lErro = objRelOpcoes.IncluirParametro("TCOMPINI", sComprador_I)
    If lErro <> AD_BOOL_TRUE Then gError 108902
    
    lErro = objRelOpcoes.IncluirParametro("TCOMPFIM", sComprador_F)
    If lErro <> AD_BOOL_TRUE Then gError 108903
    
    lErro = objRelOpcoes.IncluirParametro("TPRODINI", sProduto_I)
    If lErro <> AD_BOOL_TRUE Then gError 108904
    
    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProduto_F)
    If lErro <> AD_BOOL_TRUE Then gError 108905
    
    sCategoria = CStr(Categoria.List(Categoria.ListIndex))
    lErro = objRelOpcoes.IncluirParametro("TCATEG", sCategoria)
    If lErro <> AD_BOOL_TRUE Then gError 108906
    
    lErro = objRelOpcoes.IncluirParametro("TITEMINI", sItem_I)
    If lErro <> AD_BOOL_TRUE Then gError 108907
    
    lErro = objRelOpcoes.IncluirParametro("TITEMFIM", sItem_F)
    If lErro <> AD_BOOL_TRUE Then gError 108908
    
    Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_NF

                sOrdenacaoPor = "NotaFiscal"

            Case ORD_POR_DATA

                sOrdenacaoPor = "Data"

            Case ORD_POR_FORNECEDOR

                sOrdenacaoPor = "Fornecedor"


            Case Else
                gError 73556

    End Select

    lErro = objRelOpcoes.IncluirParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> AD_BOOL_TRUE Then gError 73557

    sOrd = ComboOrdenacao.ListIndex
    lErro = objRelOpcoes.IncluirParametro("NORDENACAO", sOrd)
    If lErro <> AD_BOOL_TRUE Then gError 73558

    lErro = Monta_Expressao_Selecao(objRelOpcoes, iFilial_I, iFilial_F, sNumero_I, sNumero_F, sForn_I, sForn_F, sSerie_I, sSerie_F, sComprador_I, sComprador_F, sProduto_I, sProduto_F, sItem_I, sItem_F, sOrdenacaoPor, sOrd, sCategoria)
    If lErro <> SUCESSO Then gError 73559

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 73541 To 73559, 108900 To 108908

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170260)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(iFilial_I As Integer, iFilial_F As Integer, sNomeFil_I As String, sNomeFil_F As String, sSerie_I As String, sSerie_F As String, sNumero_I As String, sNumero_F As String, sForn_I As String, sForn_F As String, sComprador_I As String, sComprador_F As String, sProduto_I As String, sProduto_F As String, sItem_I As String, sItem_F As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim iProdPreenchido_I As Integer, iProdPreenchido_F As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros

    'critica Codigo da Filial Inicial e Final
    If FilialEmpresaDe.List(FilialEmpresaDe.ListIndex) <> "" Then
        iFilial_I = Codigo_Extrai(FilialEmpresaDe.List(FilialEmpresaDe.ListIndex))
    Else
        iFilial_I = 0
    End If

    If FilialEmpresaAte.List(FilialEmpresaAte.ListIndex) <> "" Then
        iFilial_F = Codigo_Extrai(FilialEmpresaAte.List(FilialEmpresaAte.ListIndex))
    Else
        iFilial_F = 0
    End If

    If iFilial_I <> 0 And iFilial_F <> 0 Then

        If iFilial_I > iFilial_F Then gError 73560

    End If

    'critica NumeroNF Inicial e Final
    If NumeroDe.Text <> "" Then
        sNumero_I = CStr(NumeroDe.Text)
    Else
        sNumero_I = ""
    End If

    If NumeroAte.Text <> "" Then
        sNumero_F = CStr(NumeroAte.Text)
    Else
        sNumero_F = ""
    End If

    If sNumero_I <> "" And sNumero_F <> "" Then

        If CLng(sNumero_I) > CLng(sNumero_F) Then gError 73562

    End If

    'data inicial não pode ser maior que a final
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then

         If CDate(DataDe.Text) > CDate(DataAte.Text) Then gError 73563

    End If
    
    'critica Fornecedor Inicial e Final
    If FornecedorDe.Text <> "" Then
        sForn_I = CStr(FornecedorDe.Text)
    Else
        sForn_I = ""
    End If

    If FornecedorAte.Text <> "" Then
        sForn_F = CStr(FornecedorAte.Text)
    Else
        sForn_F = ""
    End If

    If sForn_I <> "" And sForn_F <> "" Then

        If CLng(sForn_I) > CLng(sForn_F) Then gError 73564

    End If

    'critica Serie Inicial e Final
    If SerieDe.Text <> "" Then
        sSerie_I = CStr(SerieDe.Text)
    Else
        sSerie_I = ""
    End If

    If SerieAte.Text <> "" Then
        sSerie_F = CStr(SerieAte.Text)
    Else
        sSerie_F = ""
    End If

    If sSerie_I <> "" And sSerie_F <> "" Then

        If sSerie_I > sSerie_F Then gError 73565

    End If
    
    'critica Comprador Inicial e Final
    If CompradorDe.List(CompradorDe.ListIndex) <> "" Then
        sComprador_I = CompradorDe.List(CompradorDe.ListIndex)
    Else
        sComprador_I = ""
    End If

    If CompradorAte.List(CompradorAte.ListIndex) <> "" Then
        sComprador_F = CompradorAte.List(CompradorAte.ListIndex)
    Else
        sComprador_F = ""
    End If

    If sComprador_I <> "" And sComprador_F <> "" Then

        If sComprador_I > sComprador_F Then gError 108890

    End If
    
    'data inicial de atualizacao não pode ser maior que a final
    If Trim(AtualizacaoDe.ClipText) <> "" And Trim(AtualizacaoAte.ClipText) <> "" Then

         If CDate(AtualizacaoDe.Text) > CDate(AtualizacaoAte.Text) Then gError 108891

    End If
    
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoDe.Text, sProduto_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 108892

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProduto_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoAte.Text, sProduto_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 108893

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProduto_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProduto_I > sProduto_F Then gError 108894

    End If
    
    'Critica os itens da Categoria
    If ItemDe.List(ItemDe.ListIndex) <> "" Then sItem_I = ItemDe.List(ItemDe.ListIndex)
    
    If ItemAte.List(ItemAte.ListIndex) <> "" Then sItem_F = ItemAte.List(ItemAte.ListIndex)
    
    'Se estão preenchidos => Inicial não pode ser maior do que Final
    If sItem_I <> "" And sItem_F <> "" Then
        
        If sItem_I > sItem_F Then gError 108920
        
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 73560
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            FilialEmpresaDe.SetFocus

        Case 73561
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            FilialEmpresaDe.SetFocus

        Case 73562
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMNF_INICIAL_MAIOR", gErr)
            NumeroDe.SetFocus

        Case 72563
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus

        Case 73564
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", gErr)
            FornecedorDe.SetFocus

        Case 73565
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_INICIAL_MAIOR", gErr)
            SerieDe.SetFocus
            
        Case 108890
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_INICIAL_MAIOR", gErr)
            CompradorDe.SetFocus
            
        Case 108891
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            AtualizacaoDe.SetFocus
            
        Case 108892, 108893
        
        Case 108894
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoDe.SetFocus
            
        Case 108920
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_ITEM_INICIAL_MAIOR", gErr)
            ItemDe.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170261)

    End Select

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, iFilial_I As Integer, iFilial_F As Integer, sNumero_I As String, sNumero_F As String, sForn_I As String, sForn_F As String, sSerie_I As String, sSerie_F As String, sComprador_I As String, sComprador_F As String, sProduto_I As String, sProduto_F As String, sItem_I As String, sItem_F As String, sOrdenacaoPor As String, sOrd As String, sCategoria As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao


   If iFilial_I <> 0 Then sExpressao = "S01"

   If iFilial_F <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S02"

    End If

    If sNumero_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S03"

    End If

    If sNumero_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S04"

    End If

    If Trim(DataDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S05"

    End If

    If Trim(DataAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S06"

    End If

    If sSerie_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S07"

    End If

    If sSerie_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S08"

    End If

    If sForn_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S09"

    End If

    If sForn_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S10"

    End If

    If sComprador_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S11"

    End If

    If sComprador_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S12"

    End If

    If sProduto_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S13"

    End If

    If sProduto_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S14"

    End If

    If sItem_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S15"

    End If

    If sItem_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S16"

    End If
    
        If Trim(AtualizacaoDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S17"

    End If

    If Trim(AtualizacaoAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S18"

    End If
    
    If sCategoria <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S19"

    End If

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170262)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long, iTipoOrd As Integer, iAscendente As Integer
Dim sParam As String
Dim sOrdenacaoPor As String
Dim iIndice As Integer
Dim sProdutoMascarado As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 73566

    'pega Codigo inicial
    lErro = objRelOpcoes.ObterParametro("NCODFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 73567
    
    For iIndice = 0 To FilialEmpresaDe.ListCount - 1
        If Codigo_Extrai(FilialEmpresaDe.List(iIndice)) = StrParaInt(sParam) Then
            FilialEmpresaDe.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'pega Codigo Filial
    lErro = objRelOpcoes.ObterParametro("NCODFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 73568

    For iIndice = 0 To FilialEmpresaAte.ListCount - 1
        If Codigo_Extrai(FilialEmpresaAte.List(iIndice)) = StrParaInt(sParam) Then
            FilialEmpresaAte.ListIndex = iIndice
            Exit For
        End If
    Next


    'pega  Numero inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NNOTAFISCALINIC", sParam)
    If lErro <> SUCESSO Then gError 73571

    NumeroDe.Text = sParam

    'pega numero final e exibe
    lErro = objRelOpcoes.ObterParametro("NNOTAFISCALFIM", sParam)
    If lErro <> SUCESSO Then gError 73572

    NumeroAte.Text = sParam

    'pega Fornecedor Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFORNINIC", sParam)
    If lErro <> SUCESSO Then gError 73573

    FornecedorDe.Text = sParam

    'pega Fornecedor Final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFORNFIM", sParam)
    If lErro <> SUCESSO Then gError 73574

    FornecedorAte.Text = sParam

    'pega  SErie Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TSERIEINIC", sParam)
    If lErro <> SUCESSO Then gError 73575

    SerieDe.Text = sParam
    Call SerieDe_Validate(bSGECancelDummy)

    'pega serie Final e exibe
    lErro = objRelOpcoes.ObterParametro("TSERIEFIM", sParam)
    If lErro <> SUCESSO Then gError 73576

    SerieAte.Text = sParam
    Call SerieAte_Validate(bSGECancelDummy)

    'pega data  inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAINIC", sParam)
    If lErro <> SUCESSO Then gError 73577

    Call DateParaMasked(DataDe, CDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAFIM", sParam)
    If lErro <> SUCESSO Then gError 73578

    Call DateParaMasked(DataAte, CDate(sParam))

    lErro = objRelOpcoes.ObterParametro("NITENS", sParam)
    If lErro <> SUCESSO Then gError 73579

    If sParam = "1" Then
        CheckItens.Value = 1
    Else
        CheckItens.Value = 0
    End If

    'pega data  inicial de atualizacao e exibe
    lErro = objRelOpcoes.ObterParametro("DTATUINI", sParam)
    If lErro <> SUCESSO Then gError 108910

    Call DateParaMasked(AtualizacaoDe, CDate(sParam))

    'pega data final de atualizacao e exibe
    lErro = objRelOpcoes.ObterParametro("DTATUFIM", sParam)
    If lErro <> SUCESSO Then gError 108911

    Call DateParaMasked(AtualizacaoAte, CDate(sParam))
    
    'Pega o comprador Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCOMPINI", sParam)
    If lErro <> SUCESSO Then gError 108912
        
    Call CF("SCombo_Seleciona2", CompradorDe, sParam)
    
    'Pega o comprador Final e exibe
    lErro = objRelOpcoes.ObterParametro("TCOMPFIM", sParam)
    If lErro <> SUCESSO Then gError 108913
    
    Call CF("SCombo_Seleciona2", CompradorAte, sParam)
    
    'pega  codigo do produto inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINI", sParam)
    If lErro <> SUCESSO Then gError 108914
    
    ProdutoDe.PromptInclude = False
    ProdutoDe.Text = sParam
    ProdutoDe.PromptInclude = True
    
    Call ProdutoDe_Validate(bSGECancelDummy)
    
    'pega  codigo do produto final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 108915
    
    ProdutoAte.PromptInclude = False
    ProdutoAte.Text = sParam
    ProdutoAte.PromptInclude = True
    
    Call ProdutoAte_Validate(bSGECancelDummy)
    
    'pega  a categoria e exibe
    lErro = objRelOpcoes.ObterParametro("TCATEG", sParam)
    If lErro <> SUCESSO Then gError 108916
    
    Call CF("SCombo_Seleciona2", Categoria, sParam)
    
    'pega o item inicial da categoria e exibe
    lErro = objRelOpcoes.ObterParametro("TITEMINI", sParam)
    If lErro <> SUCESSO Then gError 108917
    
    Call CF("SCombo_Seleciona2", ItemDe, sParam)
    
    'pega o item final da categoria e exibe
    lErro = objRelOpcoes.ObterParametro("TITEMFIM", sParam)
    If lErro <> SUCESSO Then gError 108918
    
    Call CF("SCombo_Seleciona2", ItemAte, sParam)
    
    lErro = objRelOpcoes.ObterParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> SUCESSO Then gError 73580

    Select Case sOrdenacaoPor

            Case "Data"

                ComboOrdenacao.ListIndex = ORD_POR_DATA

            Case "Fornecedor"

                ComboOrdenacao.ListIndex = ORD_POR_FORNECEDOR

            Case "NF"

                ComboOrdenacao.ListIndex = ORD_POR_NF

            Case Else
                gError 73581

    End Select

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 73566 To 73581, 108910 To 108918

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170263)

    End Select

    Exit Function

End Function

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

''    Parent.HelpContextID = IDH_RELOP_REQ
    Set Form_Load_Ocx = Me
    Caption = "Relação de Notas Fiscais de Pedidos de Compra"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpNFPedCompras"

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

        If Me.ActiveControl Is NumeroDe Then
            Call LabelNumeroDe_Click

        ElseIf Me.ActiveControl Is NumeroAte Then
            Call LabelNumeroAte_Click

        ElseIf Me.ActiveControl Is FornecedorDe Then
            Call LabelFornecedorDe_Click

        ElseIf Me.ActiveControl Is FornecedorAte Then
            Call LabelFornecedorAte_Click

        ElseIf Me.ActiveControl Is SerieDe Then
            Call LabelSerieDe_Click

        ElseIf Me.ActiveControl Is SerieAte Then
            Call LabelSerieAte_Click
            
        ElseIf Me.ActiveControl Is ProdutoDe Then
            Call LabelProdutoDe_Click
            
        ElseIf Me.ActiveControl Is ProdutoAte Then
            Call LabelProdutoAte_Click

        End If

    End If

End Sub

Private Sub LabelCodFilialDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodFilialDe, Source, X, Y)
End Sub

Private Sub LabelCodFilialDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodFilialDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodFilialAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodFilialAte, Source, X, Y)
End Sub

Private Sub LabelCodFilialAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodFilialAte, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Function Carrega_Serie() As Long
'Carrega a combo de Séries com as séries lidas do BD

Dim lErro As Long
Dim colSerie As New colSerie
Dim objSerie As ClassSerie

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError 73590

    'Carrega na combo
    For Each objSerie In colSerie
        SerieDe.AddItem objSerie.sSerie
        SerieAte.AddItem objSerie.sSerie
    Next

    Carrega_Serie = SUCESSO

    Exit Function

Erro_Carrega_Serie:

    Carrega_Serie = gErr

    Select Case gErr

        Case 73590

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170264)

    End Select

    Exit Function

End Function

Private Sub LabelSerieAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelSerieAte, Source, X, Y)
End Sub

Private Sub LabelSerieAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelSerieAte, Button, Shift, X, Y)
End Sub

Private Sub LabelSerieDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelSerieDe, Source, X, Y)
End Sub

Private Sub LabelSerieDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelSerieDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNumeroDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumeroDe, Source, X, Y)
End Sub

Private Sub LabelNumeroDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumeroDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNumeroAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumeroAte, Source, X, Y)
End Sub

Private Sub LabelNumeroAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumeroAte, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LabelFornecedorDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornecedorDe, Source, X, Y)
End Sub

Private Sub LabelFornecedorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornecedorDe, Button, Shift, X, Y)
End Sub

Private Sub LabelFornecedorAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornecedorAte, Source, X, Y)
End Sub

Private Sub LabelFornecedorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornecedorAte, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub ProdutoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_ProdutoDe_Validate

    If Len(Trim(ProdutoDe.ClipText)) > 0 Then
    
        sProdutoFormatado = String(STRING_PRODUTO, 0)
        
        lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 108862
        
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 108863
                
        If lErro = 28030 Then gError 108864
        
    End If
    
    Exit Sub
    
Erro_ProdutoDe_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 108862, 108863
        
        Case 108864
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170265)
            
    End Select
    
End Sub

Private Sub ProdutoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_ProdutoAte_Validate

    If Len(Trim(ProdutoAte.ClipText)) > 0 Then
        
        lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
        
        If lErro <> SUCESSO Then gError 108865
        
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 108866
        
        If lErro = 28030 Then gError 108867
        
    End If
    
    Exit Sub
    
Erro_ProdutoAte_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 108865, 108866
        
        Case 108867
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170266)
            
    End Select
    
End Sub

Private Sub LabelProdutoDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_LabelProdutoDe_Click
    
    If Len(Trim(ProdutoDe.Text)) > 0 Then
        
        'Preenche com o Produto da tela
        lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 108868
        
        objProduto.sCodigo = sProdutoFormatado
        
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProdutoDe)

   Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 108868
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170267)

    End Select

End Sub

Private Sub LabelProdutoAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_LabelProdutoAte_Click
    
    If Len(Trim(ProdutoAte.Text)) > 0 Then
        
        'Preenche com o Produto da tela
        lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 108869
        
        objProduto.sCodigo = sProdutoFormatado
        
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProdutoAte)

   Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 108869
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170268)

    End Select

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 108870
    
    ProdutoAte.Text = sProdutoMascarado

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr
    
        Case 108870
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170269)
            
    End Select
    
    Exit Sub
    
End Sub
Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 108871
    
    ProdutoDe.Text = sProdutoMascarado

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr
    
        Case 108871
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170270)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub AtualizacaoDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(AtualizacaoDe, iAlterado)
    
End Sub

Private Sub AtualizacaoAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(AtualizacaoAte, iAlterado)
    
End Sub

Private Sub ProdutoDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(ProdutoDe, iAlterado)
    
End Sub

Private Sub ProdutoAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(ProdutoAte, iAlterado)
    
End Sub

Private Sub AtualizacaoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AtualizacaoDe_Validate

    'Verifica se a AtualizacaoDe está preenchida
    If Len(Trim(AtualizacaoDe.Text)) = 0 Then Exit Sub

    'Critica a AtualizacaoDe informada
    lErro = Data_Critica(AtualizacaoDe.Text)
    If lErro <> SUCESSO Then gError 108872

    Exit Sub

Erro_AtualizacaoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 108872
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170271)

    End Select

    Exit Sub

End Sub

Private Sub AtualizacaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AtualizacaoAte_Validate

    'Verifica se a AtualizacaoDe está preenchida
    If Len(Trim(AtualizacaoAte.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(AtualizacaoAte.Text)
    If lErro <> SUCESSO Then gError 108873

    Exit Sub

Erro_AtualizacaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 108873
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170272)

    End Select

End Sub

Private Function Carrega_FilialEmpresa() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objFilialEmpresa As New AdmFiliais
Dim colFiliais As New Collection

On Error GoTo Erro_Carrega_FilialEmpresa

    'Faz a Leitura das Filiais
    lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
    If lErro <> SUCESSO Then gError 108874
    
    FilialEmpresaDe.AddItem ("")
    FilialEmpresaAte.AddItem ("")
    
    'Carrega as combos
    For Each objFilialEmpresa In colFiliais
        
        'Se nao for a EMPRESA_TODA
        If objFilialEmpresa.iCodFilial <> EMPRESA_TODA Then
            
            FilialEmpresaDe.AddItem (objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome)
            FilialEmpresaAte.AddItem (objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome)
            
        End If
        
    Next

    Carrega_FilialEmpresa = SUCESSO
    
    Exit Function
    
Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = gErr
    
    Select Case gErr
    
        Case 108874

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170273)
    
    End Select

End Function

Private Function Carrega_Compradores() As Long

Dim lErro As Long
Dim objComprador As New ClassComprador
Dim colComprador As New Collection

On Error GoTo Erro_Carrega_Compradores

    'Le os compradores
    lErro = CF("Comprador_Le_Todos", colComprador)
    If lErro <> SUCESSO And lErro <> 50126 Then gError 108875
    
    'Se nao encontrou => Erro
    If lErro = 50126 Then gError 108876
    
    CompradorDe.AddItem ("")
    CompradorAte.AddItem ("")
    
    'Carrega as combos de Compradores
    For Each objComprador In colComprador
    
        CompradorDe.AddItem objComprador.sCodUsuario
        CompradorAte.AddItem objComprador.sCodUsuario
    
    Next
    
    Carrega_Compradores = SUCESSO
    
    Exit Function
    
Erro_Carrega_Compradores:

    Carrega_Compradores = gErr
    
    Select Case gErr
    
        Case 108875
        
        Case 108876
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_NAO_CADASTRADO2", gErr)
            '??? Não existe comprador cadastrado.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170274)
    
    End Select

End Function

Private Function Carrega_Categorias() As Long

Dim lErro As Long
Dim objCategoria As New ClassCategoriaProduto
Dim colCategorias As New Collection

On Error GoTo Erro_Carrega_Categorias
    
    'Le a categoria
    lErro = CF("CategoriasProduto_Le_Todas", colCategorias)
    If lErro <> SUCESSO And lErro <> 22542 Then gError 108877
    
    'Se nao encontrou => Erro
    If lErro = 22542 Then gError 108878
    
    Categoria.AddItem ("")
    
    'Carrega as combos de Categorias
    For Each objCategoria In colCategorias
    
        Categoria.AddItem objCategoria.sCategoria
        
    Next
    
    Carrega_Categorias = SUCESSO
    
    Exit Function
    
Erro_Carrega_Categorias:

    Carrega_Categorias = gErr
    
    Select Case gErr
    
        Case 108877
        
        Case 108878
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_NAO_CADASTRADA", gErr)
            '??? Não existe categoria de produto cadastrada.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170275)
    
    End Select

End Function

Private Sub UpDownAtualizacaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownAtualizacaoAte_DownClick

    'Diminui um dia em AtualizacaoAte
    lErro = Data_Up_Down_Click(AtualizacaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 73529

    Exit Sub

Erro_UpDownAtualizacaoAte_DownClick:

    Select Case gErr

        Case 73529
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170276)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAtualizacaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownAtualizacaoAte_UpClick

    'Diminui um dia em AtualizacaoAte
    lErro = Data_Up_Down_Click(AtualizacaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 73530

    Exit Sub

Erro_UpDownAtualizacaoAte_UpClick:

    Select Case gErr

        Case 73530
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170277)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAtualizacaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownAtualizacaoDe_DownClick

    'Diminui um dia em AtualizacaoDe
    lErro = Data_Up_Down_Click(AtualizacaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 73531

    Exit Sub

Erro_UpDownAtualizacaoDe_DownClick:

    Select Case gErr

        Case 73531
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170278)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAtualizacaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownAtualizacaoDe_UpClick

    'Diminui um dia em AtualizacaoDe
    lErro = Data_Up_Down_Click(AtualizacaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 73532

    Exit Sub

Erro_UpDownAtualizacaoDe_UpClick:

    Select Case gErr

        Case 73532
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170279)

    End Select

    Exit Sub

End Sub

