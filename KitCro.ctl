VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl Kit 
   ClientHeight    =   7365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11070
   KeyPreview      =   -1  'True
   ScaleHeight     =   7365
   ScaleMode       =   0  'User
   ScaleWidth      =   10966.22
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8370
      ScaleHeight     =   495
      ScaleWidth      =   2535
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   90
      Width           =   2595
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2055
         Picture         =   "KitCro.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1554
         Picture         =   "KitCro.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1056
         Picture         =   "KitCro.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   558
         Picture         =   "KitCro.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   60
         Picture         =   "KitCro.ctx":0994
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Imprimir"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Kit"
      Height          =   1485
      Left            =   105
      TabIndex        =   35
      Top             =   645
      Width           =   10860
      Begin VB.CheckBox VersaoFormPreco 
         Caption         =   "Usar para Formação de Preços"
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
         Left            =   7815
         TabIndex        =   40
         Top             =   990
         Width           =   3000
      End
      Begin VB.CommandButton BotaoKits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   9240
         Picture         =   "KitCro.ctx":0A96
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   150
         Width           =   1500
      End
      Begin VB.ComboBox Situacao 
         Height          =   315
         ItemData        =   "KitCro.ctx":1B18
         Left            =   9240
         List            =   "KitCro.ctx":1B1A
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   645
         Width           =   1530
      End
      Begin VB.CommandButton BotaoRoteiros 
         Caption         =   "&Roteiros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7665
         Picture         =   "KitCro.ctx":1B1C
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Abre a tela de Roteiros de Fabricação"
         Top             =   150
         Width           =   1500
      End
      Begin MSMask.MaskEdBox Observacao 
         Height          =   315
         Left            =   975
         TabIndex        =   39
         Top             =   1065
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   6915
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   645
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   315
         Left            =   975
         TabIndex        =   42
         Top             =   195
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Left            =   5775
         TabIndex        =   43
         Top             =   645
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Versao 
         Height          =   315
         Left            =   975
         TabIndex        =   44
         Top             =   645
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin VB.Label ProdutoLbl 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   210
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   50
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Descricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2490
         TabIndex        =   49
         Top             =   195
         Width           =   4665
      End
      Begin VB.Label LabelVersao 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   270
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   48
         Top             =   690
         Width           =   660
      End
      Begin VB.Label Label7 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5220
         TabIndex        =   47
         Top             =   705
         Width           =   480
      End
      Begin VB.Label LabelObeservações 
         AutoSize        =   -1  'True
         Caption         =   "Obs:"
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
         Left            =   540
         TabIndex        =   46
         Top             =   1110
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Situação:"
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
         Left            =   8385
         TabIndex        =   45
         Top             =   705
         Width           =   825
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Estrutura do Produto:"
      Height          =   5115
      Left            =   120
      TabIndex        =   33
      Top             =   2130
      Width           =   5055
      Begin MSComctlLib.TreeView EstruturaProduto 
         Height          =   4770
         Left            =   90
         TabIndex        =   34
         Top             =   225
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   8414
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   354
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Componente"
      Height          =   3960
      Left            =   5250
      TabIndex        =   9
      Top             =   3285
      Width           =   5685
      Begin VB.ComboBox TipoCarga 
         Height          =   315
         ItemData        =   "KitCro.ctx":1E26
         Left            =   675
         List            =   "KitCro.ctx":1E28
         TabIndex        =   58
         Text            =   "TipoCarga"
         Top             =   2835
         Width           =   2145
      End
      Begin VB.ComboBox GrupoPesagem 
         Height          =   315
         ItemData        =   "KitCro.ctx":1E2A
         Left            =   4725
         List            =   "KitCro.ctx":1E7F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   2835
         Width           =   885
      End
      Begin VB.CommandButton BotaoIncluir 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   210
         Picture         =   "KitCro.ctx":1ED3
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3315
         Width           =   1335
      End
      Begin VB.CommandButton BotaoRemover 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2175
         Picture         =   "KitCro.ctx":3721
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3315
         Width           =   1335
      End
      Begin VB.CommandButton BotaoAlterar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4140
         Picture         =   "KitCro.ctx":5047
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3315
         Width           =   1335
      End
      Begin VB.ComboBox SiglaUM 
         Height          =   315
         Left            =   1170
         TabIndex        =   12
         Top             =   1140
         Width           =   915
      End
      Begin VB.ComboBox Composicao 
         Height          =   315
         ItemData        =   "KitCro.ctx":696D
         Left            =   3900
         List            =   "KitCro.ctx":6977
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1560
         Width           =   1680
      End
      Begin VB.ComboBox VersaoKitComp 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "KitCro.ctx":698B
         Left            =   1170
         List            =   "KitCro.ctx":698D
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   720
         Width           =   1500
      End
      Begin MSMask.MaskEdBox ProdutoSel 
         Height          =   315
         Left            =   1170
         TabIndex        =   16
         Top             =   270
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   315
         Left            =   1170
         TabIndex        =   17
         Top             =   1560
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PercentualPerda 
         Height          =   315
         Left            =   1170
         TabIndex        =   18
         Top             =   1980
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CustoStandard 
         Height          =   315
         Left            =   3900
         TabIndex        =   19
         Top             =   1965
         Width           =   1680
         _ExtentX        =   2963
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
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label TipoCargaLabel 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   165
         TabIndex        =   60
         Top             =   2880
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Grupo de pesagem:"
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
         Left            =   3030
         TabIndex        =   59
         Top             =   2880
         Width           =   1665
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Custo Standard:"
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
         Left            =   2490
         TabIndex        =   32
         Top             =   2025
         Width           =   1380
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nível:"
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
         Left            =   615
         TabIndex        =   31
         Top             =   2460
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Composição:"
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
         Left            =   2790
         TabIndex        =   30
         Top             =   1605
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade:"
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
         Left            =   105
         TabIndex        =   29
         Top             =   1605
         Width           =   1050
      End
      Begin VB.Label ComponenteLabel 
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
         Left            =   495
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   28
         Top             =   330
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Perda:"
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
         Left            =   570
         TabIndex        =   27
         Top             =   2025
         Width           =   570
      End
      Begin VB.Label Sequencial 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3915
         TabIndex        =   26
         Top             =   2415
         Width           =   285
      End
      Begin VB.Label DescProdutoSel 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2700
         TabIndex        =   25
         Top             =   270
         Width           =   2880
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "U.M.:"
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
         Left            =   675
         TabIndex        =   24
         Top             =   1185
         Width           =   480
      End
      Begin VB.Label NomeUM 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2130
         TabIndex        =   23
         Top             =   1140
         Width           =   2040
      End
      Begin VB.Label Nivel 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1170
         TabIndex        =   22
         Top             =   2415
         Width           =   285
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Sequencial:"
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
         Left            =   2850
         TabIndex        =   21
         Top             =   2460
         Width           =   1020
      End
      Begin VB.Label LabelVersaoComp 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   495
         TabIndex        =   20
         Top             =   750
         Width           =   660
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Fatores Para Cálculo de Custo"
      Height          =   1125
      Left            =   5280
      TabIndex        =   0
      Top             =   2130
      Width           =   5685
      Begin MSMask.MaskEdBox MaoDeObra 
         Height          =   315
         Left            =   1890
         TabIndex        =   1
         Top             =   270
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Energia 
         Height          =   315
         Left            =   1890
         TabIndex        =   2
         Top             =   675
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox BPF 
         Height          =   315
         Left            =   4125
         TabIndex        =   3
         Top             =   270
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Agua 
         Height          =   315
         Left            =   4125
         TabIndex        =   4
         Top             =   675
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Mão de Obra Direta:"
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
         Left            =   120
         TabIndex        =   8
         Top             =   315
         Width           =   1740
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Energia:"
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
         Left            =   1140
         TabIndex        =   7
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Gás/BPF:"
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
         Left            =   3225
         TabIndex        =   6
         Top             =   315
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Água:"
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
         Index           =   3
         Left            =   3555
         TabIndex        =   5
         Top             =   720
         Width           =   510
      End
   End
End
Attribute VB_Name = "Kit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTKit
Attribute objCT.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
    
    Set objCT = New CTKit
    Set objCT.objUserControl = Me
    
    Set objCT.gobjInfoUsu = New CTKitVGCro
    Set objCT.gobjInfoUsu.gobjTelaUsu = New CTKitCro
    
End Sub

Private Sub Agua_Change()
     Call objCT.Agua_Change
End Sub

Private Sub Agua_GotFocus()
     Call objCT.Agua_GotFocus
End Sub

Private Sub BotaoAlterar_Click()
     Call objCT.BotaoAlterar_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Private Sub BotaoImprimir_Click()
     Call objCT.BotaoImprimir_Click
End Sub

Private Sub BotaoIncluir_Click()
     Call objCT.BotaoIncluir_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub BotaoRemover_Click()
     Call objCT.BotaoRemover_Click
End Sub

Private Sub BPF_Change()
     Call objCT.BPF_Change
End Sub

Private Sub BPF_GotFocus()
     Call objCT.BPF_GotFocus
End Sub

Private Sub Data_GotFocus()
     Call objCT.Data_GotFocus
End Sub

Private Sub Data_Validate(Cancel As Boolean)
     Call objCT.Data_Validate(Cancel)
End Sub

Private Sub Energia_Change()
     Call objCT.Energia_Change
End Sub

Private Sub Energia_GotFocus()
     Call objCT.Energia_GotFocus
End Sub

Private Sub EstruturaProduto_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.EstruturaProduto_NodeClick(Node)
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

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub MaoDeObra_Change()
     Call objCT.MaoDeObra_Change
End Sub

Private Sub MaoDeObra_GotFocus()
     Call objCT.MaoDeObra_GotFocus
End Sub

Private Sub PercentualPerda_Change()
     Call objCT.PercentualPerda_Change
End Sub

Private Sub PercentualPerda_Validate(Cancel As Boolean)
     Call objCT.PercentualPerda_Validate(Cancel)
End Sub

Private Sub Produto_GotFocus()
     Call objCT.Produto_GotFocus
End Sub

Private Sub ProdutoLbl_Click()
     Call objCT.ProdutoLbl_Click
End Sub

Private Sub ComponenteLabel_Click()
     Call objCT.ComponenteLabel_Click
End Sub

Private Sub LabelVersao_Click()
     Call objCT.LabelVersao_Click
End Sub

Private Sub Observacao_Change()
     Call objCT.Observacao_Change
End Sub

Private Sub Produto_Change()
     Call objCT.Produto_Change
End Sub

Private Sub Produto_Validate(Cancel As Boolean)
     Call objCT.Produto_Validate(Cancel)
End Sub

Private Sub ProdutoSel_Change()
     Call objCT.ProdutoSel_Change
End Sub

Private Sub Composicao_Seleciona()
     Call objCT.Composicao_Seleciona
End Sub

Private Sub ProdutoSel_GotFocus()
     Call objCT.ProdutoSel_GotFocus
End Sub

Private Sub ProdutoSel_Validate(Cancel As Boolean)
     Call objCT.ProdutoSel_Validate(Cancel)
End Sub

Private Sub CustoStandard_Change()
     Call objCT.CustoStandard_Change
End Sub

Private Sub CustoStandard_Validate(Cancel As Boolean)
     Call objCT.CustoStandard_Validate(Cancel)
End Sub

Private Sub Quantidade_Change()
     Call objCT.Quantidade_Change
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)
     Call objCT.Quantidade_Validate(Cancel)
End Sub

Private Sub SiglaUM_Change()
     Call objCT.SiglaUM_Change
End Sub

Private Sub SiglaUM_Click()
     Call objCT.SiglaUM_Click
End Sub

Private Sub SiglaUM_Validate(Cancel As Boolean)
     Call objCT.SiglaUM_Validate(Cancel)
End Sub

Function Trata_Parametros(Optional objKit As ClassKit) As Long
     Trata_Parametros = objCT.Trata_Parametros(objKit)
End Function

Private Sub Situacao_Change()
     Call objCT.Situacao_Change
End Sub

Private Sub Situacao_Click()
     Call objCT.Situacao_Click
End Sub

Private Sub Versao_Change()
     Call objCT.Versao_Change
End Sub

Private Sub UpDown1_DownClick()
     Call objCT.UpDown1_DownClick
End Sub

Private Sub UpDown1_UpClick()
     Call objCT.UpDown1_UpClick
End Sub

Private Sub BotaoKits_Click()
     Call objCT.BotaoKits_Click
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub
Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub
Private Sub Nivel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Nivel, Source, X, Y)
End Sub
Private Sub Nivel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Nivel, Button, Shift, X, Y)
End Sub
Private Sub NomeUM_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NomeUM, Source, X, Y)
End Sub
Private Sub NomeUM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NomeUM, Button, Shift, X, Y)
End Sub
Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub
Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub
Private Sub DescProdutoSel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescProdutoSel, Source, X, Y)
End Sub
Private Sub DescProdutoSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescProdutoSel, Button, Shift, X, Y)
End Sub
Private Sub Sequencial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Sequencial, Source, X, Y)
End Sub
Private Sub Sequencial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Sequencial, Button, Shift, X, Y)
End Sub
Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub
Private Sub ComponenteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ComponenteLabel, Source, X, Y)
End Sub
Private Sub ComponenteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ComponenteLabel, Button, Shift, X, Y)
End Sub
Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub
Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub
Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub
Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub
Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub
Private Sub Label35_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label35, Source, X, Y)
End Sub
Private Sub Label35_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label35, Button, Shift, X, Y)
End Sub
Private Sub ProdutoLbl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoLbl, Source, X, Y)
End Sub
Private Sub ProdutoLbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoLbl, Button, Shift, X, Y)
End Sub
Private Sub Descricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Descricao, Source, X, Y)
End Sub
Private Sub Descricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Descricao, Button, Shift, X, Y)
End Sub
Private Sub LabelVersao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVersao, Source, X, Y)
End Sub
Private Sub LabelVersao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVersao, Button, Shift, X, Y)
End Sub
Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub
Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub
Private Sub LabelObeservações_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelObeservações, Source, X, Y)
End Sub
Private Sub LabelObeservações_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelObeservações, Button, Shift, X, Y)
End Sub
Private Sub LabelVersaoComp_Click()
     Call objCT.LabelVersaoComp_Click
End Sub

Private Sub VersaoKitComp_Seleciona(sVersao As String)
     Call objCT.VersaoKitComp_Seleciona(sVersao)
End Sub

Private Sub VersaoFormPreco_Click()
     Call objCT.VersaoFormPreco_Click
End Sub

Private Sub BotaoRoteiros_Click()
     Call objCT.BotaoRoteiros_Click
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


