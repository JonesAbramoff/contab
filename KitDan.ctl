VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl Kit 
   ClientHeight    =   7680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   KeyPreview      =   -1  'True
   ScaleHeight     =   7674.518
   ScaleMode       =   0  'User
   ScaleWidth      =   10995
   Begin VB.Frame Frame4 
      Caption         =   "Fatores Para Cálculo de Custo"
      Height          =   1680
      Left            =   5220
      TabIndex        =   49
      Top             =   5850
      Width           =   5685
      Begin MSMask.MaskEdBox MaoDeObra 
         Height          =   315
         Left            =   1905
         TabIndex        =   50
         Top             =   315
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         Format          =   "#,##0.00####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Energia 
         Height          =   315
         Left            =   1905
         TabIndex        =   51
         Top             =   795
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         Format          =   "#,##0.00####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox BPF 
         Height          =   315
         Left            =   4125
         TabIndex        =   52
         Top             =   315
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         Format          =   "#,##0.00####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Agua 
         Height          =   315
         Left            =   4125
         TabIndex        =   53
         Top             =   795
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         Format          =   "#,##0.00####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PesoFator5 
         Height          =   315
         Left            =   1875
         TabIndex        =   54
         Top             =   1245
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         Format          =   "#,##0.00####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PesoFator6 
         Height          =   315
         Left            =   4110
         TabIndex        =   55
         Top             =   1245
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         Format          =   "#,##0.00####"
         PromptChar      =   " "
      End
      Begin VB.Label LabelFator 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   120
         TabIndex        =   61
         Top             =   360
         Width           =   1740
      End
      Begin VB.Label LabelFator 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   60
         TabIndex        =   60
         Top             =   840
         Width           =   1755
      End
      Begin VB.Label LabelFator 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   59
         Top             =   360
         Width           =   840
      End
      Begin VB.Label LabelFator 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   3210
         TabIndex        =   58
         Top             =   840
         Width           =   855
      End
      Begin VB.Label LabelFator 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fator 6:"
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
         Index           =   6
         Left            =   3315
         TabIndex        =   57
         Top             =   1290
         Width           =   765
      End
      Begin VB.Label LabelFator 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fator 5:"
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
         Index           =   5
         Left            =   135
         TabIndex        =   56
         Top             =   1275
         Width           =   1665
      End
   End
   Begin VB.CommandButton BotaoOS 
      Caption         =   "OSs que utilizam este kit..."
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
      Left            =   2460
      TabIndex        =   48
      Top             =   60
      Width           =   2610
   End
   Begin VB.Frame Frame1 
      Caption         =   "Componente"
      Height          =   3570
      Left            =   5220
      TabIndex        =   23
      Top             =   2115
      Width           =   5685
      Begin VB.ComboBox VersaoKitComp 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "KitDan.ctx":0000
         Left            =   1170
         List            =   "KitDan.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   720
         Width           =   1500
      End
      Begin VB.ComboBox Composicao 
         Height          =   315
         ItemData        =   "KitDan.ctx":0004
         Left            =   3900
         List            =   "KitDan.ctx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1560
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.ComboBox SiglaUM 
         Height          =   315
         Left            =   1170
         TabIndex        =   9
         Top             =   1140
         Width           =   915
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
         Left            =   4110
         Picture         =   "KitDan.ctx":0022
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3000
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
         Left            =   2145
         Picture         =   "KitDan.ctx":1948
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3000
         Width           =   1335
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
         Left            =   180
         Picture         =   "KitDan.ctx":326E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3000
         Width           =   1335
      End
      Begin MSMask.MaskEdBox ProdutoSel 
         Height          =   315
         Left            =   1170
         TabIndex        =   7
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
         TabIndex        =   10
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
         TabIndex        =   12
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
         TabIndex        =   13
         Top             =   1965
         Visible         =   0   'False
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
         TabIndex        =   44
         Top             =   750
         Width           =   660
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
         TabIndex        =   35
         Top             =   2460
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Nivel 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1170
         TabIndex        =   34
         Top             =   2415
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label NomeUM 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2130
         TabIndex        =   33
         Top             =   1140
         Width           =   2040
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
         TabIndex        =   32
         Top             =   1185
         Width           =   480
      End
      Begin VB.Label DescProdutoSel 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2700
         TabIndex        =   31
         Top             =   270
         Width           =   2880
      End
      Begin VB.Label Sequencial 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3915
         TabIndex        =   30
         Top             =   2415
         Visible         =   0   'False
         Width           =   285
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
         TabIndex        =   29
         Top             =   2025
         Width           =   570
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
         TabIndex        =   27
         Top             =   1605
         Width           =   1050
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
         TabIndex        =   26
         Top             =   1605
         Visible         =   0   'False
         Width           =   1095
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
         TabIndex        =   25
         Top             =   2460
         Visible         =   0   'False
         Width           =   540
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
         TabIndex        =   24
         Top             =   2025
         Visible         =   0   'False
         Width           =   1380
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Estrutura do Produto:"
      Height          =   5475
      Left            =   90
      TabIndex        =   43
      Top             =   2070
      Width           =   5055
      Begin MSComctlLib.TreeView EstruturaProduto 
         Height          =   5115
         Left            =   90
         TabIndex        =   6
         Top             =   225
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   9022
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
   Begin VB.Frame Frame2 
      Caption         =   "Kit"
      Height          =   1485
      Left            =   75
      TabIndex        =   36
      Top             =   585
      Width           =   10860
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
         Picture         =   "KitDan.ctx":4ABC
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Abre a tela de Roteiros de Fabricação"
         Top             =   150
         Width           =   1500
      End
      Begin VB.ComboBox Situacao 
         Height          =   315
         ItemData        =   "KitDan.ctx":4DC6
         Left            =   9240
         List            =   "KitDan.ctx":4DC8
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   645
         Width           =   1530
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
         Picture         =   "KitDan.ctx":4DCA
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   150
         Width           =   1500
      End
      Begin MSMask.MaskEdBox Observacao 
         Height          =   315
         Left            =   975
         TabIndex        =   5
         Top             =   1065
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         PromptChar      =   " "
      End
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
         TabIndex        =   45
         Top             =   990
         Visible         =   0   'False
         Width           =   3000
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   4755
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   615
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   315
         Left            =   975
         TabIndex        =   0
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
         Left            =   3615
         TabIndex        =   2
         Top             =   615
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
         TabIndex        =   1
         Top             =   630
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   10
         PromptChar      =   " "
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
         TabIndex        =   42
         Top             =   705
         Width           =   825
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
         TabIndex        =   41
         Top             =   1110
         Width           =   405
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Revisão:"
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
         TabIndex        =   40
         Top             =   660
         Width           =   750
      End
      Begin VB.Label LabelVersao 
         AutoSize        =   -1  'True
         Caption         =   "Desenho:"
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
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   39
         Top             =   690
         Width           =   825
      End
      Begin VB.Label Descricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2490
         TabIndex        =   38
         Top             =   195
         Width           =   4665
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
         TabIndex        =   37
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8340
      ScaleHeight     =   495
      ScaleWidth      =   2535
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   30
      Width           =   2595
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   60
         Picture         =   "KitDan.ctx":5E4C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Imprimir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   558
         Picture         =   "KitDan.ctx":5F4E
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1056
         Picture         =   "KitDan.ctx":60A8
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1554
         Picture         =   "KitDan.ctx":6232
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2055
         Picture         =   "KitDan.ctx":6764
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
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

Private Sub BotaoOS_Click()
    Call objCT.gobjInfoUsu.gobjTelaUsu.BotaoOS_Click(objCT)
End Sub

Private Sub UserControl_Initialize()
    
    Set objCT = New CTKit
    Set objCT.objUserControl = Me

    Set objCT.gobjInfoUsu = New CTKitVGDan
    Set objCT.gobjInfoUsu.gobjTelaUsu = New CTKitDan

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

