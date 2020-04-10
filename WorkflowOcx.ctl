VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.UserControl WorkflowOcx 
   ClientHeight    =   6960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9495
   KeyPreview      =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   9495
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3930
      Index           =   1
      Left            =   135
      TabIndex        =   16
      Top             =   1485
      Width           =   9270
      Begin VB.ComboBox ValidoPara 
         Height          =   315
         ItemData        =   "WorkflowOcx.ctx":0000
         Left            =   1980
         List            =   "WorkflowOcx.ctx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   570
         Width           =   1650
      End
      Begin VB.TextBox RelAnexoGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   6195
         TabIndex        =   80
         Top             =   2160
         Width           =   285
      End
      Begin VB.TextBox RelSelGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   5895
         TabIndex        =   79
         Top             =   2145
         Width           =   285
      End
      Begin VB.TextBox RelPorEmailGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   5535
         TabIndex        =   78
         Top             =   2145
         Width           =   285
      End
      Begin VB.TextBox BrowseOpcao 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   7815
         TabIndex        =   71
         Top             =   585
         Width           =   285
      End
      Begin VB.TextBox RelModulo 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         TabIndex        =   70
         Top             =   285
         Width           =   285
      End
      Begin VB.TextBox RelOpcao 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   6000
         TabIndex        =   69
         Top             =   510
         Width           =   285
      End
      Begin VB.TextBox BrowseNome 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   7320
         TabIndex        =   68
         Top             =   510
         Width           =   285
      End
      Begin VB.TextBox BrowseModulo 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   6660
         TabIndex        =   67
         Top             =   570
         Width           =   285
      End
      Begin VB.TextBox RelNome 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   5625
         TabIndex        =   66
         Top             =   450
         Width           =   285
      End
      Begin VB.CheckBox Relatorio 
         Enabled         =   0   'False
         Height          =   210
         Left            =   4665
         TabIndex        =   65
         Top             =   705
         Width           =   855
      End
      Begin VB.CheckBox Browser 
         Enabled         =   0   'False
         Height          =   210
         Left            =   3480
         TabIndex        =   64
         Top             =   945
         Width           =   795
      End
      Begin VB.TextBox AvisoUsuGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   4950
         TabIndex        =   47
         Top             =   2130
         Width           =   285
      End
      Begin VB.TextBox LogMsgGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   4260
         TabIndex        =   46
         Top             =   2205
         Width           =   285
      End
      Begin VB.TextBox LogDocGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   3630
         TabIndex        =   45
         Top             =   2175
         Width           =   285
      End
      Begin VB.TextBox EmailMsgGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2940
         TabIndex        =   44
         Top             =   2160
         Width           =   285
      End
      Begin VB.TextBox EmailAssuntoGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2190
         TabIndex        =   43
         Top             =   2145
         Width           =   285
      End
      Begin VB.TextBox EmailParaGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1350
         TabIndex        =   42
         Top             =   2160
         Width           =   285
      End
      Begin VB.TextBox AvisoMsgGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   510
         TabIndex        =   41
         Top             =   2175
         Width           =   285
      End
      Begin VB.CheckBox Log 
         Enabled         =   0   'False
         Height          =   210
         Left            =   8235
         TabIndex        =   21
         Top             =   1800
         Width           =   600
      End
      Begin VB.CheckBox Aviso 
         Enabled         =   0   'False
         Height          =   210
         Left            =   7425
         TabIndex        =   20
         Top             =   1830
         Width           =   600
      End
      Begin VB.TextBox Regra 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   495
         TabIndex        =   19
         Top             =   1275
         Width           =   3615
      End
      Begin VB.ComboBox TipoBloqueio 
         Height          =   315
         ItemData        =   "WorkflowOcx.ctx":0033
         Left            =   5940
         List            =   "WorkflowOcx.ctx":0035
         TabIndex        =   18
         Top             =   1185
         Width           =   1650
      End
      Begin VB.CheckBox Email 
         Enabled         =   0   'False
         Height          =   210
         Left            =   8325
         TabIndex        =   17
         Top             =   1290
         Width           =   600
      End
      Begin MSFlexGridLib.MSFlexGrid GridRegras 
         Height          =   2520
         Left            =   60
         TabIndex        =   23
         Top             =   105
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   4445
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3960
      Index           =   5
      Left            =   120
      TabIndex        =   49
      Top             =   1440
      Visible         =   0   'False
      Width           =   9270
      Begin VB.Frame Frame3 
         Caption         =   "Consulta"
         Height          =   1560
         Left            =   300
         TabIndex        =   55
         Top             =   2340
         Width           =   8715
         Begin VB.ComboBox ComboOpcaoBrowser 
            Height          =   315
            ItemData        =   "WorkflowOcx.ctx":0037
            Left            =   1560
            List            =   "WorkflowOcx.ctx":0039
            Sorted          =   -1  'True
            TabIndex        =   62
            Top             =   1125
            Width           =   2730
         End
         Begin VB.ComboBox ComboModuloBrowser 
            Height          =   315
            ItemData        =   "WorkflowOcx.ctx":003B
            Left            =   1560
            List            =   "WorkflowOcx.ctx":003D
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   255
            Width           =   1905
         End
         Begin VB.ComboBox CodBrowser 
            Height          =   315
            Left            =   1560
            TabIndex        =   56
            Top             =   675
            Width           =   4275
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
            Index           =   6
            Left            =   840
            TabIndex        =   63
            Top             =   1185
            Width           =   615
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Módulo:"
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
            Left            =   780
            TabIndex        =   59
            Top             =   285
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Consulta:"
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
            Left            =   660
            TabIndex        =   58
            Top             =   720
            Width           =   810
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Relatório"
         Height          =   2145
         Left            =   270
         TabIndex        =   50
         Top             =   120
         Width           =   8715
         Begin VB.TextBox RelAnexo 
            Height          =   285
            Left            =   975
            MaxLength       =   250
            TabIndex        =   76
            Top             =   1725
            Width           =   7515
         End
         Begin VB.TextBox RelSel 
            Height          =   285
            Left            =   975
            MaxLength       =   250
            TabIndex        =   74
            Top             =   1260
            Width           =   7515
         End
         Begin VB.CheckBox RelPorEmail 
            Caption         =   "Por e-mail"
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
            Left            =   4155
            TabIndex        =   73
            Top             =   705
            Width           =   1845
         End
         Begin VB.ComboBox ComboOpcoes 
            Height          =   315
            ItemData        =   "WorkflowOcx.ctx":003F
            Left            =   1005
            List            =   "WorkflowOcx.ctx":0041
            Sorted          =   -1  'True
            TabIndex        =   60
            Top             =   750
            Width           =   3015
         End
         Begin VB.ComboBox CodRelatorio 
            Height          =   315
            Left            =   3990
            TabIndex        =   52
            Top             =   240
            Width           =   4530
         End
         Begin VB.ComboBox ComboModulo 
            Height          =   315
            ItemData        =   "WorkflowOcx.ctx":0043
            Left            =   1020
            List            =   "WorkflowOcx.ctx":0045
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   240
            Width           =   1905
         End
         Begin VB.Label Label2 
            Caption         =   "Anexo:"
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
            Left            =   270
            TabIndex        =   77
            Top             =   1770
            Width           =   585
         End
         Begin VB.Label Label2 
            Caption         =   "Filtro:"
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
            Left            =   390
            TabIndex        =   75
            Top             =   1305
            Width           =   585
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
            Index           =   5
            Left            =   270
            TabIndex        =   61
            Top             =   810
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Relatório:"
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
            Left            =   3060
            TabIndex        =   54
            Top             =   285
            Width           =   825
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Módulo:"
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
            Left            =   225
            TabIndex        =   53
            Top             =   270
            Width           =   690
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3960
      Index           =   2
      Left            =   135
      TabIndex        =   24
      Top             =   1440
      Visible         =   0   'False
      Width           =   9270
      Begin VB.TextBox EmailMsg 
         Height          =   1005
         Left            =   135
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   1785
         Width           =   9000
      End
      Begin VB.TextBox EmailAssunto 
         Height          =   285
         Left            =   915
         MaxLength       =   250
         TabIndex        =   31
         Top             =   960
         Width           =   7815
      End
      Begin VB.TextBox EmailPara 
         Height          =   255
         Left            =   915
         MaxLength       =   8000
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   375
         Width           =   7830
      End
      Begin VB.Label Label4 
         Caption         =   "Mensagem:"
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
         TabIndex        =   35
         Top             =   1530
         Width           =   1080
      End
      Begin VB.Label Label2 
         Caption         =   "Assunto:"
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
         Left            =   105
         TabIndex        =   34
         Top             =   975
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Para:"
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
         Left            =   375
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   33
         Top             =   390
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3915
      Index           =   3
      Left            =   165
      TabIndex        =   36
      Top             =   1485
      Visible         =   0   'False
      Width           =   9270
      Begin VB.TextBox AvisoMsg 
         Height          =   1005
         Left            =   135
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Top             =   330
         Width           =   9000
      End
      Begin VB.ListBox AvisoUsu 
         Columns         =   5
         Height          =   1635
         ItemData        =   "WorkflowOcx.ctx":0047
         Left            =   180
         List            =   "WorkflowOcx.ctx":0049
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   37
         Top             =   1635
         Width           =   8985
      End
      Begin VB.Label Label7 
         Caption         =   "Mensagem:"
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
         Left            =   150
         TabIndex        =   40
         Top             =   60
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   " Usuários"
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
         Left            =   90
         TabIndex        =   38
         Top             =   1410
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   3960
      Index           =   4
      Left            =   150
      TabIndex        =   25
      Top             =   1440
      Visible         =   0   'False
      Width           =   9270
      Begin VB.TextBox LogMsg 
         Height          =   930
         Left            =   195
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   1545
         Width           =   9000
      End
      Begin VB.TextBox LogDoc 
         Height          =   255
         Left            =   1485
         MaxLength       =   8000
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   405
         Width           =   3180
      End
      Begin VB.Label Label3 
         Caption         =   "Mensagem:"
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
         TabIndex        =   29
         Top             =   1305
         Width           =   1080
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   210
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   28
         Top             =   435
         Width           =   1035
      End
   End
   Begin VB.CheckBox CheckBox_Workflow_Ativo 
      Caption         =   "Workflow Ativo"
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
      Left            =   7485
      TabIndex        =   72
      Top             =   780
      Value           =   1  'Checked
      Width           =   1755
   End
   Begin VB.CheckBox Checkbox_Verifica_Sintaxe 
      Caption         =   "Verifica Sintaxe ao Sair do Campo"
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
      Left            =   6150
      TabIndex        =   48
      Top             =   1110
      Value           =   1  'Checked
      Width           =   3285
   End
   Begin VB.ComboBox Mnemonicos 
      Height          =   315
      Left            =   150
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   5895
      Width           =   3675
   End
   Begin VB.ComboBox Funcoes 
      Height          =   315
      ItemData        =   "WorkflowOcx.ctx":004B
      Left            =   3990
      List            =   "WorkflowOcx.ctx":004D
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   5895
      Width           =   3795
   End
   Begin VB.ComboBox Operadores 
      Height          =   315
      Left            =   7950
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   5895
      Width           =   1150
   End
   Begin VB.TextBox Descricao 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   540
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   6300
      Width           =   8955
   End
   Begin VB.ComboBox Transacao 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   675
      Width           =   5745
   End
   Begin VB.ComboBox Modulo 
      Height          =   315
      Left            =   1305
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   195
      Width           =   3630
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   7170
      ScaleHeight     =   450
      ScaleWidth      =   2115
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   2175
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1605
         Picture         =   "WorkflowOcx.ctx":004F
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   1095
         Picture         =   "WorkflowOcx.ctx":01CD
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   345
         Left            =   570
         Picture         =   "WorkflowOcx.ctx":06FF
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   345
         Left            =   90
         Picture         =   "WorkflowOcx.ctx":0889
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4395
      Left            =   75
      TabIndex        =   22
      Top             =   1110
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7752
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Regras"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "E-mail"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Aviso"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Log"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Relatório/Consulta"
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Operadores:"
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
      Left            =   7965
      TabIndex        =   15
      Top             =   5640
      Width           =   1050
   End
   Begin VB.Label LabelFuncoes 
      AutoSize        =   -1  'True
      Caption         =   "Funções:"
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
      Left            =   4020
      TabIndex        =   14
      Top             =   5640
      Width           =   795
   End
   Begin VB.Label LabelMnemonicos 
      AutoSize        =   -1  'True
      Caption         =   "Mnemônicos:"
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
      TabIndex        =   13
      Top             =   5640
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Transação:"
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
      Left            =   210
      TabIndex        =   8
      Top             =   750
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Módulo:"
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
      Left            =   465
      TabIndex        =   7
      Top             =   240
      Width           =   690
   End
End
Attribute VB_Name = "WorkflowOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private gobjTela As Object

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Const KEYCODE_VERIFICAR_SINTAXE = vbKeyF5

Dim iGrid_Regra_Col As Integer
Dim iGrid_ValidoPara_Col As Integer
Dim iGrid_TipoBloqueio_Col As Integer
Dim iGrid_Email_Col As Integer
Dim iGrid_Aviso_Col As Integer
Dim iGrid_Log_Col As Integer
Dim iGrid_Rel_Col As Integer
Dim iGrid_Browse_Col As Integer
Dim iGrid_EmailParaGrid_Col As Integer
Dim iGrid_EmailAssuntoGrid_Col As Integer
Dim iGrid_EmailMsgGrid_Col As Integer
Dim iGrid_AvisoMsgGrid_Col As Integer
Dim iGrid_AvisoUsuGrid_Col As Integer
Dim iGrid_LogDocGrid_Col As Integer
Dim iGrid_LogMsgGrid_Col As Integer
Dim iGrid_RelModulo_Col As Integer
Dim iGrid_RelNome_Col As Integer
Dim iGrid_RelOpcao_Col As Integer
Dim iGrid_RelPorEmail_Col As Integer
Dim iGrid_RelSel_Col As Integer
Dim iGrid_RelAnexo_Col As Integer
Dim iGrid_BrowseModulo_Col As Integer
Dim iGrid_BrowseNome_Col As Integer
Dim iGrid_BrowseOpcao_Col As Integer


Dim objGridRegra As AdmGrid
Dim iAlterado As Integer
Dim iFrameAtual As Integer
Dim objCampoAtual As Object

Dim giGridRefresh As Integer

Dim sModulo As String
Dim iTransacao As Integer

Const CONTABILIZACAO_OBRIGATORIA = 1
Const CONTABILIZACAO_NAO_OBRIGATORIA = 0

Const TAB_REGRAS = 1
Const TAB_EMAIL = 2
Const TAB_AVISO = 3
Const TAB_LOG = 4
Const TAB_RELATORIO = 5

Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

Private Sub AvisoMsg_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AvisoMsg_GotFocus()
    Set objCampoAtual = AvisoMsg
End Sub

Private Sub AvisoMsg_Validate(Cancel As Boolean)
        
Dim lErro As Long
Dim colMnemonico As New Collection
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_AvisoMsg_Validate

    GridRegras.TextMatrix(GridRegras.Row, iGrid_AvisoMsgGrid_Col) = AvisoMsg.Text

    If Checkbox_Verifica_Sintaxe.Value = MARCADO Then

        lErro = CF("MnemonicoWFW_Le", gcolModulo.Sigla(sModulo), iTransacao, colMnemonico)
        If lErro <> SUCESSO Then gError 178136

        lErro = CF("Valida_Formula_WFW", AvisoMsg.Text, TIPO_TEXTO, iInicio, iTamanho, colMnemonico)
        If lErro <> SUCESSO Then gError 178107
            
    End If

    Exit Sub
    
Erro_AvisoMsg_Validate:

    Cancel = True

    Select Case gErr

        Case 178107
            AvisoMsg.SelStart = iInicio
            AvisoMsg.SelLength = iTamanho
            
        Case 178136
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178127)
            
    End Select
    
    Exit Sub

End Sub

Private Sub AvisoUsu_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AvisoUsu_GotFocus()
    Set objCampoAtual = AvisoUsu
End Sub

Private Sub AvisoUsu_Validate(Cancel As Boolean)

Dim iIndice As Integer
Dim iLinha As Integer
    
On Error GoTo Erro_AvisoUsu_Validate
    
    GridRegras.TextMatrix(GridRegras.Row, iGrid_AvisoUsuGrid_Col) = ""
    
    For iIndice = 0 To AvisoUsu.ListCount - 1
        If AvisoUsu.Selected(iIndice) = True Then
            GridRegras.TextMatrix(GridRegras.Row, iGrid_AvisoUsuGrid_Col) = GridRegras.TextMatrix(GridRegras.Row, iGrid_AvisoUsuGrid_Col) & AvisoUsu.List(iIndice) & " "
        End If
    Next
    
    Exit Sub
    
Erro_AvisoUsu_Validate:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178128)
            
    End Select

End Sub

Private Sub CodBrowser_Click()

Dim lErro As Long
Dim colOpcoes As New Collection
Static sCodBrowser As String
Dim vOpcao As Variant
    
On Error GoTo Erro_CodBrowser_Click
    
    If Len(Trim(CodBrowser.Text)) > 0 Then

        ComboOpcaoBrowser.Clear
        
        sCodBrowser = CodBrowser.Text

        lErro = CF("BrowseOpcaoOrdenacao_Le_Opcoes", sCodBrowser, colOpcoes)
        If lErro <> SUCESSO Then gError 178395
        
        For Each vOpcao In colOpcoes
            ComboOpcaoBrowser.AddItem vOpcao
        Next
    
    End If
    
    Exit Sub

Erro_CodBrowser_Click:

    Select Case gErr

        Case 178386
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178396)
            
    End Select
    
    Exit Sub

End Sub

Private Sub CodBrowser_Validate(Cancel As Boolean)

Dim lErro As Long
Dim colOpcoes As New Collection
Static sCodBrowser As String
Dim vOpcao As Variant
    
On Error GoTo Erro_CodBrowser_Validate
    
    If Len(Trim(CodBrowser.Text)) > 0 And CodBrowser.Text <> sCodBrowser Then

        ComboOpcaoBrowser.Clear
        
        sCodBrowser = CodBrowser.Text

        lErro = CF("BrowseOpcaoOrdenacao_Le_Opcoes", sCodBrowser, colOpcoes)
        If lErro <> SUCESSO Then gError 178386
        
        For Each vOpcao In colOpcoes
            ComboOpcaoBrowser.AddItem vOpcao
        Next
    
    End If
    
    GridRegras.TextMatrix(GridRegras.Row, iGrid_BrowseNome_Col) = CodBrowser.Text
    GridRegras.TextMatrix(GridRegras.Row, iGrid_BrowseOpcao_Col) = ComboOpcaoBrowser.Text
    
    Exit Sub

Erro_CodBrowser_Validate:

    Cancel = True

    Select Case gErr

        Case 178386
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178387)
            
    End Select
    
    Exit Sub

End Sub

Private Sub CodRelatorio_Click()

Dim colRelParametros As New Collection
Dim lErro As Long
Dim objRelOpcoesAux As AdmRelOpcoes
Static sCodRel As String

On Error GoTo Erro_CodRelatorio_Click

    If Len(Trim(CodRelatorio.Text)) > 0 Then

        sCodRel = CodRelatorio.Text

        ComboOpcoes.Clear
        
        'le os nomes das opcoes do relatório existentes no BD
        lErro = CF("RelOpcoes_Le_Todos", sCodRel, colRelParametros)
        If lErro <> SUCESSO Then gError 178393
    
        'preenche o ComboBox com os nomes das opções do relatório
        For Each objRelOpcoesAux In colRelParametros
            ComboOpcoes.AddItem objRelOpcoesAux.sNome
        Next

    End If

    Exit Sub

Erro_CodRelatorio_Click:

    Select Case gErr

        Case 178393
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178394)
            
    End Select
    
    Exit Sub

End Sub

Private Sub CodRelatorio_Validate(Cancel As Boolean)

Dim colRelParametros As New Collection
Dim lErro As Long
Dim objRelOpcoesAux As AdmRelOpcoes
Static sCodRel As String

On Error GoTo Erro_CodRelatorio_Validate

    If Len(Trim(CodRelatorio.Text)) > 0 And CodRelatorio.Text <> sCodRel Then

        sCodRel = CodRelatorio.Text

        ComboOpcoes.Clear
        
        'le os nomes das opcoes do relatório existentes no BD
        lErro = CF("RelOpcoes_Le_Todos", sCodRel, colRelParametros)
        If lErro <> SUCESSO Then gError 178384
    
        'preenche o ComboBox com os nomes das opções do relatório
        For Each objRelOpcoesAux In colRelParametros
            ComboOpcoes.AddItem objRelOpcoesAux.sNome
        Next

    End If

    GridRegras.TextMatrix(GridRegras.Row, iGrid_RelNome_Col) = CodRelatorio.Text
    GridRegras.TextMatrix(GridRegras.Row, iGrid_RelOpcao_Col) = ComboOpcoes.Text
    GridRegras.TextMatrix(GridRegras.Row, iGrid_RelPorEmail_Col) = CStr(RelPorEmail.Value)
    GridRegras.TextMatrix(GridRegras.Row, iGrid_RelSel_Col) = RelSel.Text
    GridRegras.TextMatrix(GridRegras.Row, iGrid_RelAnexo_Col) = RelAnexo.Text

    Exit Sub

Erro_CodRelatorio_Validate:

    Cancel = True

    Select Case gErr

        Case 178384
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178385)
            
    End Select
    
    Exit Sub

End Sub

Private Sub ComboModulo_Validate(Cancel As Boolean)

        GridRegras.TextMatrix(GridRegras.Row, iGrid_RelModulo_Col) = ComboModulo.Text
        GridRegras.TextMatrix(GridRegras.Row, iGrid_RelNome_Col) = CodRelatorio.Text
        GridRegras.TextMatrix(GridRegras.Row, iGrid_RelOpcao_Col) = ComboOpcoes.Text
        GridRegras.TextMatrix(GridRegras.Row, iGrid_RelPorEmail_Col) = RelPorEmail.Value
        GridRegras.TextMatrix(GridRegras.Row, iGrid_RelSel_Col) = RelSel.Text
        GridRegras.TextMatrix(GridRegras.Row, iGrid_RelAnexo_Col) = RelAnexo.Text
    
End Sub

Private Sub ComboModuloBrowser_Validate(Cancel As Boolean)
        GridRegras.TextMatrix(GridRegras.Row, iGrid_BrowseModulo_Col) = ComboModuloBrowser.Text
        GridRegras.TextMatrix(GridRegras.Row, iGrid_BrowseNome_Col) = CodBrowser.Text
        GridRegras.TextMatrix(GridRegras.Row, iGrid_BrowseOpcao_Col) = ComboOpcaoBrowser.Text
End Sub

Private Sub ComboOpcaoBrowser_Validate(Cancel As Boolean)

Dim iIndice As Integer

    GridRegras.TextMatrix(GridRegras.Row, iGrid_BrowseOpcao_Col) = ""

    For iIndice = 0 To ComboOpcaoBrowser.ListCount - 1
        If ComboOpcaoBrowser.List(iIndice) = ComboOpcaoBrowser.Text Then
            GridRegras.TextMatrix(GridRegras.Row, iGrid_BrowseOpcao_Col) = ComboOpcaoBrowser.Text
            Exit For
        End If
    Next

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

Dim iIndice As Integer

    GridRegras.TextMatrix(GridRegras.Row, iGrid_RelOpcao_Col) = ""

    For iIndice = 0 To ComboOpcoes.ListCount - 1
        If ComboOpcoes.List(iIndice) = ComboOpcoes.Text Then
            GridRegras.TextMatrix(GridRegras.Row, iGrid_RelOpcao_Col) = ComboOpcoes.Text
            Exit For
        End If
    Next
    
End Sub

Private Sub EmailAssunto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub EmailAssunto_GotFocus()
    Set objCampoAtual = EmailAssunto
End Sub

Private Sub EmailAssunto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim colMnemonico As New Collection
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_EmailAssunto_Validate

    GridRegras.TextMatrix(GridRegras.Row, iGrid_EmailAssuntoGrid_Col) = EmailAssunto.Text
    
    If Checkbox_Verifica_Sintaxe.Value = MARCADO Then

        lErro = CF("MnemonicoWFW_Le", gcolModulo.Sigla(sModulo), iTransacao, colMnemonico)
        If lErro <> SUCESSO Then gError 178137

        lErro = CF("Valida_Formula_WFW", EmailAssunto.Text, TIPO_TEXTO, iInicio, iTamanho, colMnemonico)
        If lErro <> SUCESSO Then gError 178108
            
    End If

    Exit Sub
    
Erro_EmailAssunto_Validate:

    Cancel = True

    Select Case gErr

        Case 178108
            EmailAssunto.SelStart = iInicio
            EmailAssunto.SelLength = iTamanho
            
        Case 178137
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178126)
            
    End Select
    
    Exit Sub

End Sub

Private Sub EmailMsg_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub EmailMsg_GotFocus()
    Set objCampoAtual = EmailMsg
End Sub

Private Sub EmailMsg_Validate(Cancel As Boolean)

Dim lErro As Long
Dim colMnemonico As New Collection
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_EmailMsg_Validate

    GridRegras.TextMatrix(GridRegras.Row, iGrid_EmailMsgGrid_Col) = EmailMsg.Text

    If Checkbox_Verifica_Sintaxe.Value = MARCADO Then

        lErro = CF("MnemonicoWFW_Le", gcolModulo.Sigla(sModulo), iTransacao, colMnemonico)
        If lErro <> SUCESSO Then gError 178138

        lErro = CF("Valida_Formula_WFW", EmailMsg.Text, TIPO_TEXTO, iInicio, iTamanho, colMnemonico)
        If lErro <> SUCESSO Then gError 178109
            
    End If

    Exit Sub
    
Erro_EmailMsg_Validate:

    Cancel = True

    Select Case gErr

        Case 178109
            EmailMsg.SelStart = iInicio
            EmailMsg.SelLength = iTamanho
            
        Case 178138
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178125)
            
    End Select
    
    Exit Sub

End Sub

Private Sub EmailPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub EmailPara_GotFocus()
    Set objCampoAtual = EmailPara
End Sub

Private Sub EmailPara_Validate(Cancel As Boolean)
    
Dim lErro As Long
Dim colMnemonico As New Collection
Dim iInicio As Integer
Dim iTamanho As Integer
    
On Error GoTo Erro_EmailPara_Validate
    
    GridRegras.TextMatrix(GridRegras.Row, iGrid_EmailParaGrid_Col) = EmailPara.Text
    
    If Checkbox_Verifica_Sintaxe.Value = MARCADO Then

        lErro = CF("MnemonicoWFW_Le", gcolModulo.Sigla(sModulo), iTransacao, colMnemonico)
        If lErro <> SUCESSO Then gError 178257

        lErro = CF("Valida_Formula_WFW", EmailPara.Text, TIPO_TEXTO, iInicio, iTamanho, colMnemonico)
        If lErro <> SUCESSO Then gError 178258
            
    End If

    Exit Sub
    
Erro_EmailPara_Validate:

    Cancel = True

    Select Case gErr

        Case 178257

        Case 178258
            EmailPara.SelStart = iInicio
            EmailPara.SelLength = iTamanho
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178130)
            
    End Select
    
    Exit Sub

End Sub

Private Sub LogDoc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LogDoc_GotFocus()
    Set objCampoAtual = LogDoc
End Sub

Private Sub LogDoc_Validate(Cancel As Boolean)

Dim lErro As Long
Dim colMnemonico As New Collection
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_LogDoc_Validate

    GridRegras.TextMatrix(GridRegras.Row, iGrid_LogDocGrid_Col) = LogDoc.Text

    If Checkbox_Verifica_Sintaxe.Value = MARCADO Then

        lErro = CF("MnemonicoWFW_Le", gcolModulo.Sigla(sModulo), iTransacao, colMnemonico)
        If lErro <> SUCESSO Then gError 178139

        lErro = CF("Valida_Formula_WFW", LogDoc.Text, TIPO_TEXTO, iInicio, iTamanho, colMnemonico)
        If lErro <> SUCESSO Then gError 178110
            
    End If

    Exit Sub
    
Erro_LogDoc_Validate:

    Cancel = True

    Select Case gErr

        Case 178110
            LogDoc.SelStart = iInicio
            LogDoc.SelLength = iTamanho
            
        Case 178139
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178124)
            
    End Select
    
    Exit Sub

End Sub

Private Sub LogMsg_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LogMsg_GotFocus()
    Set objCampoAtual = LogMsg
End Sub

Private Sub LogMsg_Validate(Cancel As Boolean)

Dim lErro As Long
Dim colMnemonico As New Collection
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_LogMsg_Validate

    GridRegras.TextMatrix(GridRegras.Row, iGrid_LogMsgGrid_Col) = LogMsg.Text

    If Checkbox_Verifica_Sintaxe.Value = MARCADO Then

        lErro = CF("MnemonicoWFW_Le", gcolModulo.Sigla(sModulo), iTransacao, colMnemonico)
        If lErro <> SUCESSO Then gError 178140

        lErro = CF("Valida_Formula_WFW", LogMsg.Text, TIPO_TEXTO, iInicio, iTamanho, colMnemonico)
        If lErro <> SUCESSO Then gError 178111
            
    End If

    Exit Sub
    
Erro_LogMsg_Validate:

    Cancel = True

    Select Case gErr

        Case 178111
            LogMsg.SelStart = iInicio
            LogMsg.SelLength = iTamanho
            
        Case 178140
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178123)
            
    End Select
    
    Exit Sub

End Sub

Public Sub Regra_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Regra_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRegra)
    
End Sub

Public Sub Regra_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegra)
    
End Sub

Public Sub Regra_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegra.objControle = Regra
    lErro = Grid_Campo_Libera_Foco(objGridRegra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub RelAnexo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub RelAnexo_GotFocus()
    Set objCampoAtual = RelAnexo
End Sub

Private Sub RelAnexo_Validate(Cancel As Boolean)
    
Dim lErro As Long
Dim colMnemonico As New Collection
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_RelAnexo_Validate

    GridRegras.TextMatrix(GridRegras.Row, iGrid_RelAnexo_Col) = RelAnexo.Text

    If Checkbox_Verifica_Sintaxe.Value = MARCADO Then

        lErro = CF("MnemonicoWFW_Le", gcolModulo.Sigla(sModulo), iTransacao, colMnemonico)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = CF("Valida_Formula_WFW", RelAnexo.Text, TIPO_TEXTO, iInicio, iTamanho, colMnemonico)
        If lErro <> SUCESSO Then gError 178111
            
    End If

    Exit Sub
    
Erro_RelAnexo_Validate:

    Cancel = True

    Select Case gErr

        Case 178111
            RelAnexo.SelStart = iInicio
            RelAnexo.SelLength = iTamanho
            
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178123)
            
    End Select
    
    Exit Sub

End Sub

Private Sub RelSel_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub RelSel_GotFocus()
    Set objCampoAtual = RelSel
End Sub

Private Sub RelSel_Validate(Cancel As Boolean)
    
Dim lErro As Long
Dim colMnemonico As New Collection
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_RelSel_Validate

    GridRegras.TextMatrix(GridRegras.Row, iGrid_RelSel_Col) = RelSel.Text

    If Checkbox_Verifica_Sintaxe.Value = MARCADO Then

        lErro = CF("MnemonicoWFW_Le", gcolModulo.Sigla(sModulo), iTransacao, colMnemonico)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = CF("Valida_Formula_WFW", RelSel.Text, TIPO_TEXTO, iInicio, iTamanho, colMnemonico)
        If lErro <> SUCESSO Then gError 178111
            
    End If

    Exit Sub
    
Erro_RelSel_Validate:

    Cancel = True

    Select Case gErr

        Case 178111
            RelSel.SelStart = iInicio
            RelSel.SelLength = iTamanho
            
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178123)
            
    End Select
    
    Exit Sub

End Sub

Private Sub TipoBloqueio_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoBloqueio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRegra)

End Sub

Private Sub TipoBloqueio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegra)

End Sub

Private Sub TipoBloqueio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegra.objControle = TipoBloqueio
    lErro = Grid_Campo_Libera_Foco(objGridRegra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGridRegra = Nothing
    Set gobjTela = Nothing
    
End Sub

Public Sub Funcoes_Click()

Dim iPos As Integer
Dim lErro As Long
Dim objFormulaFuncao As New ClassFormulaFuncao
Dim lPos As Long
Dim sFuncao As String
    
On Error GoTo Erro_Funcoes_Click
    
    objFormulaFuncao.sFuncaoCombo = Funcoes.Text
    
    'retorna os dados da funcao passada como parametro
    lErro = CF("FormulaFuncao_Le", objFormulaFuncao)
    If lErro <> SUCESSO And lErro <> 36088 Then gError 178033
    
    Descricao.Text = objFormulaFuncao.sFuncaoDesc
    
    lPos = InStr(1, Funcoes.Text, "(")
    If lPos = 0 Then
        sFuncao = Funcoes.Text
    Else
        sFuncao = Mid(Funcoes.Text, 1, lPos)
    End If
    
    lErro = Funcoes1(sFuncao)
    If lErro <> SUCESSO Then gError 178034
    
    Exit Sub
    
Erro_Funcoes_Click:

    Select Case gErr
    
        Case 178033, 178034
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178035)
            
    End Select
        
    Exit Sub

End Sub

Public Sub GridRegras_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridRegra, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRegra, iAlterado)
    End If
    
    
End Sub

Public Sub GridRegras_GotFocus()
    
    Call Grid_Recebe_Foco(objGridRegra)

End Sub

Public Sub GridRegras_EnterCell()
    
    Call Grid_Entrada_Celula(objGridRegra, iAlterado)
    
End Sub

Public Sub GridRegras_LeaveCell()
    
    Call Saida_Celula(objGridRegra)
    
End Sub

Public Sub GridRegras_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridRegra)
    
End Sub

Public Sub GridRegras_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridRegra, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRegra, iAlterado)
    End If

End Sub

Public Sub GridRegras_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridRegra)

End Sub

Public Sub GridRegras_RowColChange()

Dim iPos As Integer
Dim iPosNovo As Integer
Dim iIndice As Integer
Dim sUsuario As String

On Error GoTo Erro_GridRegras_RowColChange

    Call Grid_RowColChange(objGridRegra)
    
    If giGridRefresh = 0 Then
    
        If GridRegras.Row <= objGridRegra.iLinhasExistentes Then
            
            EmailPara.Text = GridRegras.TextMatrix(GridRegras.Row, iGrid_EmailParaGrid_Col)
            EmailAssunto.Text = GridRegras.TextMatrix(GridRegras.Row, iGrid_EmailAssuntoGrid_Col)
            EmailMsg.Text = GridRegras.TextMatrix(GridRegras.Row, iGrid_EmailMsgGrid_Col)
            AvisoMsg.Text = GridRegras.TextMatrix(GridRegras.Row, iGrid_AvisoMsgGrid_Col)
            LogDoc.Text = GridRegras.TextMatrix(GridRegras.Row, iGrid_LogDocGrid_Col)
            LogMsg.Text = GridRegras.TextMatrix(GridRegras.Row, iGrid_LogMsgGrid_Col)
            
            
            ComboModulo.ListIndex = -1
            CodRelatorio.Clear
            ComboOpcoes.Clear
            
            If Len(Trim(GridRegras.TextMatrix(GridRegras.Row, iGrid_RelModulo_Col))) > 0 Then
            
                For iIndice = 0 To ComboModulo.ListCount - 1
                    If ComboModulo.List(iIndice) = GridRegras.TextMatrix(GridRegras.Row, iGrid_RelModulo_Col) Then
                        ComboModulo.ListIndex = iIndice
                        Exit For
                    End If
                Next
            
            End If
            
            For iIndice = 0 To CodRelatorio.ListCount - 1
                If CodRelatorio.List(iIndice) = GridRegras.TextMatrix(GridRegras.Row, iGrid_RelNome_Col) Then
                    CodRelatorio.ListIndex = iIndice
                    Exit For
                End If
            Next
            
            For iIndice = 0 To ComboOpcoes.ListCount - 1
                If ComboOpcoes.List(iIndice) = GridRegras.TextMatrix(GridRegras.Row, iGrid_RelOpcao_Col) Then
                    ComboOpcoes.ListIndex = iIndice
                    Exit For
                End If
            Next
            
            RelPorEmail.Value = StrParaInt(GridRegras.TextMatrix(GridRegras.Row, iGrid_RelPorEmail_Col))
            RelSel.Text = GridRegras.TextMatrix(GridRegras.Row, iGrid_RelSel_Col)
            RelAnexo.Text = GridRegras.TextMatrix(GridRegras.Row, iGrid_RelAnexo_Col)
            
            ComboModuloBrowser.ListIndex = -1
            CodBrowser.Clear
            ComboOpcaoBrowser.Clear
            
            
            If Len(Trim(GridRegras.TextMatrix(GridRegras.Row, iGrid_BrowseModulo_Col))) > 0 Then
            
                For iIndice = 0 To ComboModuloBrowser.ListCount - 1
                    If ComboModuloBrowser.List(iIndice) = GridRegras.TextMatrix(GridRegras.Row, iGrid_BrowseModulo_Col) Then
                        ComboModuloBrowser.ListIndex = iIndice
                        Exit For
                    End If
                Next
            End If
            
            For iIndice = 0 To CodBrowser.ListCount - 1
                If CodBrowser.List(iIndice) = GridRegras.TextMatrix(GridRegras.Row, iGrid_BrowseNome_Col) Then
                    CodBrowser.ListIndex = iIndice
                    Exit For
                End If
            Next
            
            For iIndice = 0 To ComboOpcaoBrowser.ListCount - 1
                If ComboOpcaoBrowser.List(iIndice) = GridRegras.TextMatrix(GridRegras.Row, iGrid_BrowseOpcao_Col) Then
                    ComboOpcaoBrowser.ListIndex = iIndice
                    Exit For
                End If
            Next
            
            For iIndice = 0 To AvisoUsu.ListCount - 1
                AvisoUsu.Selected(iIndice) = False
            Next
            
            iPos = 1
            
            iPosNovo = InStr(iPos, GridRegras.TextMatrix(GridRegras.Row, iGrid_AvisoUsuGrid_Col), " ")
            
            Do While iPosNovo > 0
            
                sUsuario = Mid(GridRegras.TextMatrix(GridRegras.Row, iGrid_AvisoUsuGrid_Col), iPos, iPosNovo - iPos)
            
                For iIndice = 0 To AvisoUsu.ListCount - 1
                    If AvisoUsu.List(iIndice) = sUsuario Then
                        AvisoUsu.Selected(iIndice) = True
                        Exit For
                    End If
                Next
           
                iPos = iPosNovo + 1
                iPosNovo = InStr(iPos, GridRegras.TextMatrix(GridRegras.Row, iGrid_AvisoUsuGrid_Col), " ")
            Loop
            
        Else
            For iIndice = 0 To AvisoUsu.ListCount - 1
                AvisoUsu.Selected(iIndice) = False
            Next
            Call Limpa_Tela(Me)
            
        End If
    
    End If
    
    Exit Sub
    
Erro_GridRegras_RowColChange:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178081)
            
    End Select
        
    Exit Sub
       
End Sub

Public Sub GridRegras_Scroll()

    Call Grid_Scroll(objGridRegra)
    
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objCombo As Object
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Form_Load
    
    iFrameAtual = TAB_REGRAS
    
    Set objGridRegra = New AdmGrid
    
    'inicializa o grid de lancamentos padrão
    lErro = Inicializa_Grid_Regras(objGridRegra)
    If lErro <> SUCESSO Then gError 178008

    'carrega a combobox de modulos
    lErro = Carga_Combobox_Modulo()
    If lErro <> SUCESSO Then gError 178009

    'carrega a combobox de funcoes
    lErro = Carga_Combobox_Funcoes()
    If lErro <> SUCESSO Then gError 178010
    
    'carrega a combobox de operadores
    lErro = Carga_Combobox_Operadores()
    If lErro <> SUCESSO Then gError 178011
    
    lErro = Carga_ListBox_AvisoUsu()
    If lErro <> SUCESSO Then gError 178088
    
    Set objCombo = ComboModulo
    
    'carrega a combobox de modulos
    lErro = Carga_Combobox_ComboModulo(objCombo)
    If lErro <> SUCESSO Then gError 178377
    
    Set objCombo = ComboModuloBrowser
    
    'carrega a combobox de modulos
    lErro = Carga_Combobox_ComboModulo(objCombo)
    If lErro <> SUCESSO Then gError 178378
    
    objUsuarios.sCodUsuario = gsUsuario
    
    lErro = CF("Usuarios_Le", objUsuarios)
    If lErro <> SUCESSO And lErro <> 40832 Then gError 178417

    If lErro <> AD_SQL_SUCESSO Then gError 178418
    
    CheckBox_Workflow_Ativo.Value = objUsuarios.iWorkFlowAtivo
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 178008 To 178011, 178088, 178377, 178378, 178417
            
        Case 178418
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, gsUsuario)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178012)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(ByVal objTela As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
    
    Set gobjTela = objTela
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178036)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Private Function Carga_Combobox_Funcoes() As Long
'carrega a combobox que contem as funcoes disponiveis

Dim colFormulaFuncao As New Collection
Dim objFormulaFuncao As ClassFormulaFuncao
Dim lErro As Long
    
On Error GoTo Erro_Carga_Combobox_Funcoes
        
    'leitura das funcoes no BD
    lErro = CF("FormulaFuncao_Le_Todos", colFormulaFuncao)
    If lErro <> SUCESSO Then gError 178037
    
    For Each objFormulaFuncao In colFormulaFuncao
        
        Funcoes.AddItem objFormulaFuncao.sFuncaoCombo
                
    Next
    
    Carga_Combobox_Funcoes = SUCESSO

    Exit Function

Erro_Carga_Combobox_Funcoes:

    Carga_Combobox_Funcoes = gErr

    Select Case gErr

        Case 178037
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178038)

    End Select
    
    Exit Function

End Function

Private Function Carga_Combobox_Operadores() As Long
'carrega a combobox que contem os operadores disponiveis

Dim colFormulaOperador As New Collection
Dim objFormulaOperador As ClassFormulaOperador
Dim lErro As Long
    
On Error GoTo Erro_Carga_Combobox_Operadores
        
    'leitura dos operadores no BD
    lErro = CF("FormulaOperador_Le_Todos", colFormulaOperador)
    If lErro <> SUCESSO Then gError 178039
    
    For Each objFormulaOperador In colFormulaOperador
        
        Operadores.AddItem objFormulaOperador.sOperadorCombo
                
    Next
    
    Carga_Combobox_Operadores = SUCESSO

    Exit Function

Erro_Carga_Combobox_Operadores:

    Carga_Combobox_Operadores = gErr

    Select Case gErr

        Case 178039
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178040)

    End Select
    
    Exit Function

End Function

Private Function Carga_Combobox_Mnemonicos(sModulo As String, iTransacao As Integer) As Long
'carrega a combobox que contem os mnemonicos disponiveis para a transacao selecionada.

Dim colMnemonico As New Collection
Dim objMnemonico As ClassMnemonicoWFW
Dim lErro As Long
    
On Error GoTo Erro_Carga_Combobox_Mnemonicos
        
    Mnemonicos.Enabled = True
    Mnemonicos.Clear
        
    'leitura dos mnemonicos no BD para o modulo/transacao em questão
    lErro = CF("MnemonicoWFW_Le", gcolModulo.Sigla(sModulo), iTransacao, colMnemonico)
    If lErro <> SUCESSO Then gError 178031
    
    For Each objMnemonico In colMnemonico
        
        Mnemonicos.AddItem objMnemonico.sMnemonicoCombo
                
    Next
    
    Carga_Combobox_Mnemonicos = SUCESSO

    Exit Function

Erro_Carga_Combobox_Mnemonicos:

    Carga_Combobox_Mnemonicos = gErr

    Select Case gErr

        Case 178031
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178032)

    End Select
    
    Exit Function

End Function

Private Function Carga_Tipos_Bloqueio(iTransacao As Integer) As Long
'carrega a combobox que contem os tipos de bloqueio para a transacao selecionada.

Dim colTipoDeBloqueio As New Collection
Dim objTipoBloqueio As ClassTipoDeBloqueio
Dim lErro As Long
    
On Error GoTo Erro_Carga_Tipos_Bloqueio
        
    TipoBloqueio.Clear
        
    If iTransacao = TRANSACAOWFW_PEDIDO_VENDA Or iTransacao = TRANSACAOWFW_ORCAMENTO_VENDA Or iTransacao = TRANSACAOWFW_ORCAMENTO_SERVICO Then
        
        'Preenche colTipoDeBloqueio com os tipos de Bloqueio existentes na tabela TiposDeBloqueio.
        lErro = CF("TiposDeBloqueio_Le_Todos", colTipoDeBloqueio)
        If lErro <> SUCESSO Then gError 178453
        
        For Each objTipoBloqueio In colTipoDeBloqueio
    
            If objTipoBloqueio.iCodigo <> BLOQUEIO_PARCIAL And objTipoBloqueio.iCodigo <> BLOQUEIO_NAO_RESERVA And objTipoBloqueio.iCodigo <> BLOQUEIO_CREDITO And objTipoBloqueio.iCodigo <> BLOQUEIO_TOTAL And objTipoBloqueio.iCodigo <> BLOQUEIO_DIAS_ATRASO Then
                'Adiciona o item na Lista de Tabela de Preços
                TipoBloqueio.AddItem CInt(objTipoBloqueio.iCodigo) & SEPARADOR & objTipoBloqueio.sNomeReduzido
                TipoBloqueio.ItemData(TipoBloqueio.NewIndex) = objTipoBloqueio.iCodigo
            End If
        Next
    
    End If
    
    Carga_Tipos_Bloqueio = SUCESSO

    Exit Function

Erro_Carga_Tipos_Bloqueio:

    Carga_Tipos_Bloqueio = gErr

    Select Case gErr

        Case 178453
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178454)

    End Select
    
    Exit Function

End Function


Private Function Carga_Combobox_Modulo() As Long
'carrega a combobox com  os módulos disponiveis para o sistema

Dim lErro As Long
Dim iIndice As Integer
    
On Error GoTo Erro_Carga_Combobox_Modulo
        
    For iIndice = 1 To gcolModulo.Count
        If gcolModulo.Item(iIndice).iAtivo = MODULO_ATIVO Then
'            If gcolModulo.Item(iIndice).sSigla <> MODULO_ADM And gcolModulo.Item(iIndice).sSigla <> MODULO_CONTABILIDADE And gcolModulo.Item(iIndice).sSigla <> MODULO_PCP Then
                Modulo.AddItem gcolModulo.Item(iIndice).sNome
'            End If
        End If
    Next
    
    Carga_Combobox_Modulo = SUCESSO

    Exit Function

Erro_Carga_Combobox_Modulo:

    Carga_Combobox_Modulo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178041)

    End Select
    
    Exit Function

End Function

Private Function Carga_Combobox_ComboModulo(objCombo As Object) As Long
'carrega a combobox com  os módulos disponiveis para o sistema

Dim lErro As Long
Dim iIndice As Integer
    
On Error GoTo Erro_Carga_Combobox_ComboModulo
        
    For iIndice = 1 To gcolModulo.Count
        If gcolModulo.Item(iIndice).iAtivo = MODULO_ATIVO Then
            If gcolModulo.Item(iIndice).sSigla <> MODULO_ADM Then
                objCombo.AddItem gcolModulo.Item(iIndice).sNome
            End If
        End If
    Next
    
    objCombo.AddItem ""
    
    Carga_Combobox_ComboModulo = SUCESSO

    Exit Function

Erro_Carga_Combobox_ComboModulo:

    Carga_Combobox_ComboModulo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178381)

    End Select
    
    Exit Function

End Function

Private Function Carga_Combobox_Transacao(sModulo As String) As Long
'carrega a combobox que contem as transacoes disponiveis para o modulo selecionado.

Dim colTransacao As New Collection
Dim objTransacao As ClassTransacaoWFW
Dim lErro As Long
    
On Error GoTo Erro_Carga_Combobox_Transacao
        
    Transacao.Enabled = True
    Transacao.Clear
        
    'leitura das contas no BD
    lErro = CF("TransacaoWFW_Le_Todos", gcolModulo.Sigla(sModulo), colTransacao)
    If lErro <> SUCESSO Then gError 178017
    
    For Each objTransacao In colTransacao
        
        Transacao.AddItem objTransacao.sTransacaoTela
        Transacao.ItemData(Transacao.NewIndex) = objTransacao.iCodigo
    Next
    
    giGridRefresh = 1
    Call Grid_Limpa(objGridRegra)
    giGridRefresh = 0
    
    Mnemonicos.Clear
    Mnemonicos.Enabled = False
    
    Carga_Combobox_Transacao = SUCESSO

    Exit Function

Erro_Carga_Combobox_Transacao:

    Carga_Combobox_Transacao = gErr

    Select Case gErr

        Case 178017
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178018)

    End Select
    
    Exit Function

End Function

Private Function Carga_ListBox_AvisoUsu() As Long
'carrega a listbox que contem os usuarios

Dim lErro As Long
Dim colUsuarios As New Collection
Dim objUsuario As ClassUsuarios
    
On Error GoTo Erro_Carga_ListBox_AvisoUsu
        
    'leitura dos operadores no BD
    lErro = CF("Usuarios_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then gError 178089
    
    For Each objUsuario In colUsuarios
        
        AvisoUsu.AddItem objUsuario.sCodUsuario
                
    Next
    
    Carga_ListBox_AvisoUsu = SUCESSO

    Exit Function

Erro_Carga_ListBox_AvisoUsu:

    Carga_ListBox_AvisoUsu = gErr

    Select Case gErr

        Case 178089
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178090)

    End Select
    
    Exit Function

End Function

Private Function Inicializa_Grid_Regras(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Regras
    
    'tela em questão
    Set objGridRegra.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Regra")
    objGridInt.colColuna.Add ("Válido Para")
    objGridInt.colColuna.Add ("Bloqueio")
    objGridInt.colColuna.Add ("E-mail")
    objGridInt.colColuna.Add ("Aviso")
    objGridInt.colColuna.Add ("Log")
    objGridInt.colColuna.Add ("Relatório")
    objGridInt.colColuna.Add ("Consulta")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (Regra.Name)
    objGridInt.colCampo.Add (ValidoPara.Name)
    objGridInt.colCampo.Add (TipoBloqueio.Name)
    objGridInt.colCampo.Add (Email.Name)
    objGridInt.colCampo.Add (Aviso.Name)
    objGridInt.colCampo.Add (Log.Name)
    objGridInt.colCampo.Add (Relatorio.Name)
    objGridInt.colCampo.Add (Browser.Name)
    objGridInt.colCampo.Add (EmailParaGrid.Name)
    objGridInt.colCampo.Add (EmailAssuntoGrid.Name)
    objGridInt.colCampo.Add (EmailMsgGrid.Name)
    objGridInt.colCampo.Add (AvisoMsgGrid.Name)
    objGridInt.colCampo.Add (AvisoUsuGrid.Name)
    objGridInt.colCampo.Add (LogDocGrid.Name)
    objGridInt.colCampo.Add (LogMsgGrid.Name)
    
    objGridInt.colCampo.Add (RelModulo.Name)
    objGridInt.colCampo.Add (RelNome.Name)
    objGridInt.colCampo.Add (RelOpcao.Name)
    objGridInt.colCampo.Add (RelPorEmailGrid.Name)
    objGridInt.colCampo.Add (RelSelGrid.Name)
    objGridInt.colCampo.Add (RelAnexoGrid.Name)
    objGridInt.colCampo.Add (BrowseModulo.Name)
    objGridInt.colCampo.Add (BrowseNome.Name)
    objGridInt.colCampo.Add (BrowseOpcao.Name)
    
    iGrid_Regra_Col = 1
    iGrid_ValidoPara_Col = 2
    iGrid_TipoBloqueio_Col = 3
    iGrid_Email_Col = 4
    iGrid_Aviso_Col = 5
    iGrid_Log_Col = 6
    iGrid_Rel_Col = 7
    iGrid_Browse_Col = 8
    iGrid_EmailParaGrid_Col = 9
    iGrid_EmailAssuntoGrid_Col = 10
    iGrid_EmailMsgGrid_Col = 11
    iGrid_AvisoMsgGrid_Col = 12
    iGrid_AvisoUsuGrid_Col = 13
    iGrid_LogDocGrid_Col = 14
    iGrid_LogMsgGrid_Col = 15
    
    iGrid_RelModulo_Col = 16
    iGrid_RelNome_Col = 17
    iGrid_RelOpcao_Col = 18
    iGrid_RelPorEmail_Col = 19
    iGrid_RelSel_Col = 20
    iGrid_RelAnexo_Col = 21
    iGrid_BrowseModulo_Col = 22
    iGrid_BrowseNome_Col = 23
    iGrid_BrowseOpcao_Col = 24
    
    EmailParaGrid.Width = 0
    EmailAssuntoGrid.Width = 0
    EmailMsgGrid.Width = 0
    AvisoMsgGrid.Width = 0
    AvisoUsuGrid.Width = 0
    LogDocGrid.Width = 0
    LogMsgGrid.Width = 0
    RelModulo.Width = 0
    RelNome.Width = 0
    RelOpcao.Width = 0
    RelPorEmailGrid.Width = 0
    RelSelGrid.Width = 0
    RelAnexoGrid.Width = 0
    BrowseModulo.Width = 0
    BrowseNome.Width = 0
    BrowseOpcao.Width = 0
    
    objGridInt.objGrid = GridRegras
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 100 + 1
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7
        
    GridRegras.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    Call Grid_Inicializa(objGridInt)

    GridRegras.ColWidth(iGrid_EmailParaGrid_Col) = 0
    GridRegras.ColWidth(iGrid_EmailAssuntoGrid_Col) = 0
    GridRegras.ColWidth(iGrid_EmailMsgGrid_Col) = 0
    GridRegras.ColWidth(iGrid_AvisoMsgGrid_Col) = 0
    GridRegras.ColWidth(iGrid_AvisoUsuGrid_Col) = 0
    GridRegras.ColWidth(iGrid_LogDocGrid_Col) = 0
    GridRegras.ColWidth(iGrid_LogMsgGrid_Col) = 0
    
    GridRegras.ColWidth(iGrid_RelModulo_Col) = 0
    GridRegras.ColWidth(iGrid_RelNome_Col) = 0
    GridRegras.ColWidth(iGrid_RelOpcao_Col) = 0
    GridRegras.ColWidth(iGrid_RelPorEmail_Col) = 0
    GridRegras.ColWidth(iGrid_RelSel_Col) = 0
    GridRegras.ColWidth(iGrid_RelAnexo_Col) = 0
    GridRegras.ColWidth(iGrid_BrowseModulo_Col) = 0
    GridRegras.ColWidth(iGrid_BrowseNome_Col) = 0
    GridRegras.ColWidth(iGrid_BrowseOpcao_Col) = 0

    Inicializa_Grid_Regras = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_Regras:

    Inicializa_Grid_Regras = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178007)
        
    End Select

    Exit Function
        
End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then
    
        Select Case GridRegras.Col
    
            Case iGrid_Regra_Col
            
                lErro = Saida_Celula_Regra(objGridInt)
                If lErro <> SUCESSO Then gError 178051
                
            Case iGrid_TipoBloqueio_Col
            
                lErro = Saida_Celula_TipoBloqueio(objGridInt)
                If lErro <> SUCESSO Then gError 178052

            Case iGrid_ValidoPara_Col
            
                lErro = Saida_Celula_Padrao(objGridInt, ValidoPara)
                If lErro <> SUCESSO Then gError 178052

        End Select
    
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 178056
        
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula:

    Saida_Celula = gErr
    
    Select Case gErr
    
        Case 178051 To 178055
    
        Case 178056
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178057)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Regra(objGridInt As AdmGrid) As Long
'faz a critica da celula regra do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer
Dim sValor As String
Dim objHistPadrao As New ClassHistPadrao
Dim colMnemonico As New Collection

On Error GoTo Erro_Saida_Celula_Regra

    Set objGridInt.objControle = Regra

    If Len(Trim(Regra.Text)) > 0 Then
    
        If Checkbox_Verifica_Sintaxe.Value = 1 Then
    
            lErro = CF("MnemonicoWFW_Le", gcolModulo.Sigla(sModulo), iTransacao, colMnemonico)
            If lErro <> SUCESSO Then gError 178134
    
            lErro = CF("Valida_Formula_WFW", Regra.Text, TIPO_BOOLEANO, iInicio, iTamanho, colMnemonico)
            If lErro <> SUCESSO Then gError 178048
                
        End If
        
        If GridRegras.Row - GridRegras.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
        If Len(Trim(GridRegras.TextMatrix(GridRegras.Row, iGrid_ValidoPara_Col))) = 0 Then
            GridRegras.TextMatrix(GridRegras.Row, iGrid_ValidoPara_Col) = ValidoPara.List(0)
        End If
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 178049

    Saida_Celula_Regra = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Regra:

    Saida_Celula_Regra = gErr
    
    Select Case gErr
    
        Case 178048
            Regra.SelStart = iInicio
            Regra.SelLength = iTamanho
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 178049, 178134
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178050)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoBloqueio(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_TipoBloqueio

    Set objGridInt.objControle = TipoBloqueio
    
    'Verifica se o Tipo foi preenchido
    If Len(Trim(TipoBloqueio.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If TipoBloqueio.Text <> TipoBloqueio.List(TipoBloqueio.ListIndex) Then

            'Tenta selecioná-lo na combo
            lErro = Combo_Seleciona_Grid(TipoBloqueio, iCodigo)
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError 178455

            'Não foi encontrado
            If lErro = 25085 Then gError 178456
            If lErro = 25086 Then gError 178457

        End If

        
        'Acrescenta uma linha no Grid se for o caso
        If GridRegras.Row - GridRegras.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 178042

    Saida_Celula_TipoBloqueio = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoBloqueio:

    Saida_Celula_TipoBloqueio = gErr

    Select Case gErr

        Case 178042, 178455
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 178456
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_NAO_ENCONTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 178457
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_NAO_ENCONTRADO1", gErr, TipoBloqueio.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 178458)

    End Select

    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 178058
    
    Call Limpa_Tela_WFW

    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr
    
        Case 178058
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178059)
            
    End Select
    
    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long
'grava os dados da tela

Dim lErro As Long
Dim colRegraWFW As New Collection
Dim iAtivoWFW As Integer
    
On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se a transacao está preenchida. Se estiver é sinal que o modulo tambem esta
    If Transacao.ListIndex = -1 Then gError 178060
  
    'Move os dados do grid de regras para a coleção colRegraWFW
    lErro = Grid_Regra(colRegraWFW)
    If lErro <> SUCESSO Then gError 178068
    
    iAtivoWFW = CheckBox_Workflow_Ativo.Value
    
    'Grava o modelo padrão de contabilização em questão
    lErro = CF("RegraWFW_Grava", colRegraWFW, iAtivoWFW, gsUsuario)
    If lErro <> SUCESSO Then gError 178069
    
    GL_objMDIForm.MousePointer = vbDefault
    
    If iAtivoWFW = WORKFLOW_ATIVO Then
        gobjTela.Timer1.Interval = 60000
    Else
        gobjTela.Timer1.Interval = 0
    End If
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 178060
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSACAO_NAO_PREENCHIDA", gErr)
        
        Case 178068, 178069
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178070)
            
    End Select
    
    Exit Function
    
End Function

Public Sub BotaoExcluir_Click()
    
Dim lErro As Long
Dim objRegraWFW As New ClassRegraWFW
Dim vbMsgRes As VbMsgBoxResult
    
On Error GoTo Erro_BotaoExcluir_Click
     
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se a transacao está preenchida. Se estiver é sinal que o modulo tambem esta
    If Transacao.ListIndex = -1 Then gError 178097
    
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_REGRAWFW", Transacao.Text)
    
    If vbMsgRes = vbYes Then
    
        objRegraWFW.sModulo = gcolModulo.Sigla(sModulo)
        objRegraWFW.iTransacao = Transacao.ItemData(Transacao.ListIndex)
        objRegraWFW.sUsuario = gsUsuario
    
        'exclui o modelo padrão de contabilização em questão
        lErro = CF("RegraWFW_Exclui", objRegraWFW)
        If lErro <> SUCESSO Then gError 178100
    
        Call Limpa_Tela_WFW
        
        iAlterado = 0
        
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 178097
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSACAO_NAO_PREENCHIDA", gErr)
        
        Case 178100
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178101)
        
    End Select

    Exit Sub
    
End Sub

Function Limpa_Tela_WFW() As Long

    giGridRefresh = 1
    Call Grid_Limpa(objGridRegra)
    giGridRefresh = 0

    Transacao.ListIndex = -1
    Call Limpa_Tela(Me)
    iTransacao = 0
    
    ComboModulo.ListIndex = -1
    CodRelatorio.Text = ""
    ComboOpcoes.Text = ""
    ComboModuloBrowser.ListIndex = -1
    CodBrowser.Text = ""
    ComboOpcaoBrowser.Text = ""
    
    Limpa_Tela_WFW = SUCESSO
    
End Function

Function Grid_Regra(colRegraWFW As Collection) As Long
'move os dados do grid para a colecao colRegraWFW

Dim iIndice1 As Integer
Dim iIndice As Integer
Dim objRegraWFW As ClassRegraWFW
Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer
Dim colMnemonico As New Collection
Dim iPos As Integer
Dim iPosNovo As Integer
Dim sUsuario As String

On Error GoTo Erro_Grid_Regra

    lErro = CF("MnemonicoWFW_Le", gcolModulo.Sigla(sModulo), iTransacao, colMnemonico)
    If lErro <> SUCESSO Then gError 178135

    For iIndice1 = 1 To objGridRegra.iLinhasExistentes
        
        Set objRegraWFW = New ClassRegraWFW
        
        objRegraWFW.sModulo = gcolModulo.Sigla(sModulo)
        objRegraWFW.iTransacao = iTransacao
        
        objRegraWFW.iItem = iIndice1
        objRegraWFW.sUsuario = gsUsuario
            
        If Len(Trim(GridRegras.TextMatrix(iIndice1, iGrid_Regra_Col))) = 0 Then gError 178405
            
        lErro = CF("Valida_Formula_WFW", GridRegras.TextMatrix(iIndice1, iGrid_Regra_Col), TIPO_BOOLEANO, iInicio, iTamanho, colMnemonico)
        If lErro <> SUCESSO Then gError 178061
            
        objRegraWFW.sRegra = GridRegras.TextMatrix(iIndice1, iGrid_Regra_Col)
        
        If Len(GridRegras.TextMatrix(iIndice1, iGrid_TipoBloqueio_Col)) > 0 Then
            For iIndice = 0 To TipoBloqueio.ListCount - 1
                If TipoBloqueio.List(iIndice) = GridRegras.TextMatrix(iIndice1, iGrid_TipoBloqueio_Col) Then
                    objRegraWFW.iTipoBloqueio = TipoBloqueio.ItemData(iIndice)
                    Exit For
                End If
            Next
        End If
        
        If Codigo_Extrai(GridRegras.TextMatrix(iIndice1, iGrid_ValidoPara_Col)) = 2 Then
            objRegraWFW.sUsuario = ""
        End If
        
        If Len(GridRegras.TextMatrix(iIndice1, iGrid_EmailParaGrid_Col)) > 0 Then
        
            lErro = CF("Valida_Formula_WFW", GridRegras.TextMatrix(iIndice1, iGrid_EmailParaGrid_Col), TIPO_TEXTO, iInicio, iTamanho, colMnemonico)
            If lErro <> SUCESSO Then gError 178259
    
            objRegraWFW.sEmailPara = GridRegras.TextMatrix(iIndice1, iGrid_EmailParaGrid_Col)
    
        End If
        
        If Len(GridRegras.TextMatrix(iIndice1, iGrid_EmailAssuntoGrid_Col)) > 0 Then
        
            lErro = CF("Valida_Formula_WFW", GridRegras.TextMatrix(iIndice1, iGrid_EmailAssuntoGrid_Col), TIPO_TEXTO, iInicio, iTamanho, colMnemonico)
            If lErro <> SUCESSO Then gError 178062
    
            objRegraWFW.sEmailAssunto = GridRegras.TextMatrix(iIndice1, iGrid_EmailAssuntoGrid_Col)
    
        End If
            
        If Len(GridRegras.TextMatrix(iIndice1, iGrid_EmailMsgGrid_Col)) > 0 Then
        
            lErro = CF("Valida_Formula_WFW", GridRegras.TextMatrix(iIndice1, iGrid_EmailMsgGrid_Col), TIPO_TEXTO, iInicio, iTamanho, colMnemonico)
            If lErro <> SUCESSO Then gError 178063
    
            objRegraWFW.sEmailMsg = GridRegras.TextMatrix(iIndice1, iGrid_EmailMsgGrid_Col)
    
        End If
            
            
        If Len(GridRegras.TextMatrix(iIndice1, iGrid_AvisoMsgGrid_Col)) > 0 Then
        
            lErro = CF("Valida_Formula_WFW", GridRegras.TextMatrix(iIndice1, iGrid_AvisoMsgGrid_Col), TIPO_TEXTO, iInicio, iTamanho, colMnemonico)
            If lErro <> SUCESSO Then gError 178064
    
            objRegraWFW.sAvisoMsg = GridRegras.TextMatrix(iIndice1, iGrid_AvisoMsgGrid_Col)
    
        End If
            
        If Len(GridRegras.TextMatrix(iIndice1, iGrid_LogDocGrid_Col)) > 0 Then
        
            lErro = CF("Valida_Formula_WFW", GridRegras.TextMatrix(iIndice1, iGrid_LogDocGrid_Col), TIPO_TEXTO, iInicio, iTamanho, colMnemonico)
            If lErro <> SUCESSO Then gError 178065
    
            objRegraWFW.sLogDoc = GridRegras.TextMatrix(iIndice1, iGrid_LogDocGrid_Col)
    
        End If
            
        If Len(GridRegras.TextMatrix(iIndice1, iGrid_LogMsgGrid_Col)) > 0 Then
        
            lErro = CF("Valida_Formula_WFW", GridRegras.TextMatrix(iIndice1, iGrid_LogMsgGrid_Col), TIPO_TEXTO, iInicio, iTamanho, colMnemonico)
            If lErro <> SUCESSO Then gError 178066
    
            objRegraWFW.sLogMsg = GridRegras.TextMatrix(iIndice1, iGrid_LogMsgGrid_Col)
    
        End If
    
        If Len(GridRegras.TextMatrix(iIndice1, iGrid_RelSel_Col)) > 0 Then
        
            lErro = CF("Valida_Formula_WFW", GridRegras.TextMatrix(iIndice1, iGrid_RelSel_Col), TIPO_TEXTO, iInicio, iTamanho, colMnemonico)
            If lErro <> SUCESSO Then gError 201451
    
            objRegraWFW.sRelSel = GridRegras.TextMatrix(iIndice1, iGrid_RelSel_Col)
    
        End If
            
        If Len(GridRegras.TextMatrix(iIndice1, iGrid_RelAnexo_Col)) > 0 Then
        
            lErro = CF("Valida_Formula_WFW", GridRegras.TextMatrix(iIndice1, iGrid_RelAnexo_Col), TIPO_TEXTO, iInicio, iTamanho, colMnemonico)
            If lErro <> SUCESSO Then gError 201452
    
            objRegraWFW.sRelAnexo = GridRegras.TextMatrix(iIndice1, iGrid_RelAnexo_Col)
    
        End If
            
        iPos = 1
        
        iPosNovo = InStr(iPos, GridRegras.TextMatrix(iIndice1, iGrid_AvisoUsuGrid_Col), " ")
        
        Do While iPosNovo > 0
        
            sUsuario = Mid(GridRegras.TextMatrix(iIndice1, iGrid_AvisoUsuGrid_Col), iPos, iPosNovo - iPos)
        
            objRegraWFW.colUsuarios.Add sUsuario
            
            iPos = iPosNovo + 1
            iPosNovo = InStr(iPos, GridRegras.TextMatrix(iIndice1, iGrid_AvisoUsuGrid_Col), " ")
        Loop
        
        If Len(Trim(GridRegras.TextMatrix(iIndice1, iGrid_RelModulo_Col))) > 0 Then
            objRegraWFW.sRelModulo = gcolModulo.Sigla(GridRegras.TextMatrix(iIndice1, iGrid_RelModulo_Col))
        End If
        objRegraWFW.sRelNome = GridRegras.TextMatrix(iIndice1, iGrid_RelNome_Col)
        objRegraWFW.sRelOpcao = GridRegras.TextMatrix(iIndice1, iGrid_RelOpcao_Col)
        objRegraWFW.iRelPorEmail = StrParaInt(GridRegras.TextMatrix(iIndice1, iGrid_RelPorEmail_Col))
        objRegraWFW.sRelSel = GridRegras.TextMatrix(iIndice1, iGrid_RelSel_Col)
        objRegraWFW.sRelAnexo = GridRegras.TextMatrix(iIndice1, iGrid_RelAnexo_Col)
        
        If Len(Trim(GridRegras.TextMatrix(iIndice1, iGrid_BrowseModulo_Col))) > 0 Then
            objRegraWFW.sBrowseModulo = gcolModulo.Sigla(GridRegras.TextMatrix(iIndice1, iGrid_BrowseModulo_Col))
        End If
        objRegraWFW.sBrowseNome = GridRegras.TextMatrix(iIndice1, iGrid_BrowseNome_Col)
        objRegraWFW.sBrowseOpcao = GridRegras.TextMatrix(iIndice1, iGrid_BrowseOpcao_Col)
        
        'Armazena o objeto objRegraWFW na coleção colRegraWFW
        colRegraWFW.Add objRegraWFW
        
    Next
    
    Grid_Regra = SUCESSO

    Exit Function

Erro_Grid_Regra:

    Grid_Regra = gErr

    Select Case gErr
    
        Case 178061
            GridRegras.Col = iGrid_Regra_Col
            TabStrip1.Tabs.Item(TAB_REGRAS).Selected = True
            GridRegras.SetFocus
            
        Case 178062
            TabStrip1.Tabs.Item(TAB_EMAIL).Selected = True
            EmailAssunto.SetFocus
            EmailAssunto.SelStart = iInicio
            EmailAssunto.SelLength = iTamanho
            
        Case 178063
            TabStrip1.Tabs.Item(TAB_EMAIL).Selected = True
            EmailMsg.SetFocus
            EmailMsg.SelStart = iInicio
            EmailMsg.SelLength = iTamanho
            
        Case 178064
            TabStrip1.Tabs.Item(TAB_AVISO).Selected = True
            AvisoMsg.SetFocus
            AvisoMsg.SelStart = iInicio
            AvisoMsg.SelLength = iTamanho
            
        Case 178065
            TabStrip1.Tabs.Item(TAB_LOG).Selected = True
            LogDoc.SetFocus
            LogDoc.SelStart = iInicio
            LogDoc.SelLength = iTamanho
            
        Case 178066
            TabStrip1.Tabs.Item(TAB_LOG).Selected = True
            LogMsg.SetFocus
            LogMsg.SelStart = iInicio
            LogMsg.SelLength = iTamanho
            
        Case 178135
            
        Case 178259
            TabStrip1.Tabs.Item(TAB_EMAIL).Selected = True
            EmailPara.SetFocus
            EmailPara.SelStart = iInicio
            EmailPara.SelLength = iTamanho
            
        Case 201451
            TabStrip1.Tabs.Item(TAB_RELATORIO).Selected = True
            RelSel.SetFocus
            RelSel.SelStart = iInicio
            RelSel.SelLength = iTamanho
            
        Case 201452
            TabStrip1.Tabs.Item(TAB_RELATORIO).Selected = True
            RelAnexo.SetFocus
            RelAnexo.SelStart = iInicio
            RelAnexo.SelLength = iTamanho
            
        Case 178405
            Call Rotina_Erro(vbOKOnly, "REGRAWFW_NAO_PREENCHIDA", gErr, iIndice1)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178067)
            
    End Select
    
    Exit Function

End Function

Public Sub BotaoLimpar_Click()

Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim lDoc As Long
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 178102

    Call Limpa_Tela_WFW
    
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 178102
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178103)
        
    End Select
    
End Sub

Public Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Public Sub Mnemonicos_Click()

Dim iPos As Integer
Dim lErro As Long
Dim lPos As Long
Dim objMnemonico As New ClassMnemonicoWFW
Dim sMnemoncico As String
Dim sMnemonico As String

On Error GoTo Erro_Mnemonicos_Click
    
    If Len(Mnemonicos.Text) > 0 Then
    
        objMnemonico.sModulo = gcolModulo.Sigla(sModulo)
        objMnemonico.iTransacao = iTransacao
        objMnemonico.sMnemonicoCombo = Mnemonicos.Text
    
        'retorna os dados do mnemonico passado como parametro
        lErro = CF("MnemonicoWFW_Le_Mnemonico", objMnemonico)
        If lErro <> SUCESSO And lErro <> 178118 Then gError 178120
        
        If lErro = 178118 Then gError 178119
        
        Descricao.Text = objMnemonico.sMnemonicoDesc
        
        lPos = InStr(1, Mnemonicos.Text, "(")
        If lPos = 0 Then
            sMnemonico = Mnemonicos.Text
        Else
            sMnemonico = Mid(Mnemonicos.Text, 1, lPos)
        End If
        
        lErro = Mnemonicos1(sMnemonico)
        If lErro <> SUCESSO Then gError 178121
        
    End If
    
    Exit Sub
    
Erro_Mnemonicos_Click:

    Select Case gErr
    
        Case 178119
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_INEXISTENTE", Err, objMnemonico.sMnemonicoCombo)
    
        Case 178120, 178121
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178122)
            
    End Select
        
    Exit Sub
        
End Sub

Public Sub Modelo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Modulo_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Modulo_Click

    If Modulo.ListIndex = -1 Then Exit Sub

    If Modulo.Text = sModulo Then Exit Sub

    'verifica se existe a necessidade de salvar o modelo antigo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 178020
    
    'pega o valor do novo Modulo
    sModulo = Modulo.Text

    'carrega a combobox de transações com as transações referentes ao modulo em questão
    lErro = Carga_Combobox_Transacao(sModulo)
    If lErro <> SUCESSO Then gError 178021

    GridRegras.TopRow = 1
    
    iAlterado = 0

    Exit Sub

Erro_Modulo_Click:

    Select Case gErr

        Case 178020
            For iIndice = 0 To Modulo.ListCount - 1
                If Modulo.List(iIndice) = sModulo Then
                    Modulo.ListIndex = iIndice
                    Exit For
                End If
            Next

        Case 178021
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178022)
            
    End Select
        
    Exit Sub
    
End Sub

Public Sub Operadores_Click()

Dim iPos As Integer
Dim lErro As Long
Dim objFormulaOperador As New ClassFormulaOperador
Dim lPos As Integer

On Error GoTo Erro_Operadores_Click
    
    objFormulaOperador.sOperadorCombo = Operadores.Text
    
    'retorna os dados do operador passado como parametro
    lErro = CF("FormulaOperador_Le", objFormulaOperador)
    If lErro <> SUCESSO And lErro <> 36098 Then gError 178112
    
    Descricao.Text = objFormulaOperador.sOperadorDesc
    
    Call Operadores1
    
    Exit Sub
    
Erro_Operadores_Click:

    Select Case gErr
    
        Case 178112
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178113)
            
    End Select
        
    Exit Sub

End Sub

Public Sub TabStrip1_Click()
    
Dim iLinha As Integer
    
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
        
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        iFrameAtual = TabStrip1.SelectedItem.Index
        
        If TabStrip1.SelectedItem.Index = TAB_REGRAS And GridRegras.Row > 0 Then
        
            iLinha = GridRegras.Row
            
            If Len(Trim(AvisoMsg.Text)) > 0 Or Len(Trim(GridRegras.TextMatrix(GridRegras.Row, iGrid_AvisoUsuGrid_Col))) > 0 Then
                GridRegras.TextMatrix(GridRegras.Row, iGrid_Aviso_Col) = MARCADO
            Else
                GridRegras.TextMatrix(GridRegras.Row, iGrid_Aviso_Col) = DESMARCADO
            End If
            
            If Len(Trim(EmailPara.Text)) > 0 Or Len(Trim(EmailAssunto.Text)) > 0 Or Len(Trim(EmailMsg.Text)) > 0 Then
                GridRegras.TextMatrix(GridRegras.Row, iGrid_Email_Col) = MARCADO
            Else
                GridRegras.TextMatrix(GridRegras.Row, iGrid_Email_Col) = DESMARCADO
            End If
            
            If Len(Trim(LogDoc.Text)) > 0 Or Len(Trim(LogMsg.Text)) > 0 Then
                GridRegras.TextMatrix(GridRegras.Row, iGrid_Log_Col) = MARCADO
            Else
                GridRegras.TextMatrix(GridRegras.Row, iGrid_Log_Col) = DESMARCADO
            End If
            
            If Len(Trim(GridRegras.TextMatrix(GridRegras.Row, iGrid_RelModulo_Col))) > 0 Or Len(Trim(GridRegras.TextMatrix(GridRegras.Row, iGrid_RelNome_Col))) > 0 Or Len(Trim(GridRegras.TextMatrix(GridRegras.Row, iGrid_RelOpcao_Col))) > 0 Then
                GridRegras.TextMatrix(GridRegras.Row, iGrid_Rel_Col) = MARCADO
            Else
                GridRegras.TextMatrix(GridRegras.Row, iGrid_Rel_Col) = DESMARCADO
            End If
            
        
            If Len(Trim(GridRegras.TextMatrix(GridRegras.Row, iGrid_BrowseModulo_Col))) > 0 Or Len(Trim(GridRegras.TextMatrix(GridRegras.Row, iGrid_BrowseNome_Col))) > 0 Or Len(Trim(GridRegras.TextMatrix(GridRegras.Row, iGrid_BrowseOpcao_Col))) > 0 Then
                GridRegras.TextMatrix(GridRegras.Row, iGrid_Browse_Col) = MARCADO
            Else
                GridRegras.TextMatrix(GridRegras.Row, iGrid_Browse_Col) = DESMARCADO
            End If
            
            giGridRefresh = 1
            Call Grid_Refresh_Checkbox(objGridRegra)
            giGridRefresh = 0
            GridRegras.Row = iLinha
            
        End If
        
    End If
    
End Sub

Public Sub Transacao_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objTransacaoCTB As New ClassTransacaoCTB

On Error GoTo Erro_Transacao_Click

    If Transacao.ListIndex = -1 Then Exit Sub

    If Transacao.ItemData(Transacao.ListIndex) = iTransacao Then Exit Sub

    'verifica se existe a necessidade de salvar o modelo antigo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 178023
    
    'pega o valor do novo Modulo
    iTransacao = Transacao.ItemData(Transacao.ListIndex)

    'carrega a combobox de mnemonicos com os mnemonicos referentes ao modulo/transacao em questão
    lErro = Carga_Combobox_Mnemonicos(sModulo, iTransacao)
    If lErro <> SUCESSO Then gError 178024

    lErro = Carga_Tipos_Bloqueio(iTransacao)
    If lErro <> SUCESSO Then gError 178452

    'carrega o grid com os dados do modelo em questão
    lErro = Carga_Grid(sModulo, iTransacao)
    If lErro <> SUCESSO Then gError 178071
    
    GridRegras.TopRow = 1
    
    iAlterado = 0

    Exit Sub

Erro_Transacao_Click:

    Select Case gErr

        Case 178023
            For iIndice = 0 To Transacao.ListCount - 1
                If Transacao.ItemData(iIndice) = iTransacao Then
                    Transacao.ListIndex = iIndice
                    Exit For
                End If
            Next

        Case 178024, 178071, 178452
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178025)
            
    End Select
        
    Exit Sub

End Sub
   
Function Carga_Grid(ByVal sModulo As String, ByVal iTransacao As Integer) As Long
'carrega o grid com os dados do transacao em questão

Dim lErro As Long
Dim colRegraWFW As New Collection
Dim objRegraWFW As New ClassRegraWFW
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim sUsuario As String
    
On Error GoTo Erro_Carga_Grid
        
    giGridRefresh = 1
    Call Grid_Limpa(objGridRegra)
    giGridRefresh = 0
        
    TipoBloqueio.Enabled = False
    
    If iTransacao = TRANSACAOWFW_PEDIDO_COMPRA Then
    
        TipoBloqueio.Enabled = True
    
    ElseIf iTransacao = TRANSACAOWFW_PEDIDO_VENDA Or iTransacao = TRANSACAOWFW_ORCAMENTO_VENDA Or iTransacao = TRANSACAOWFW_ORCAMENTO_SERVICO Then
        
        TipoBloqueio.Enabled = True
        
    End If
        
    objRegraWFW.sModulo = gcolModulo.Sigla(sModulo)
    objRegraWFW.iTransacao = iTransacao
    objRegraWFW.sUsuario = gsUsuario
        
    'leitura das regras do modelo/transacao em questão
    lErro = CF("RegraWFW_Le_Transacao", objRegraWFW, colRegraWFW)
    If lErro <> SUCESSO Then gError 178072
    
    For Each objRegraWFW In colRegraWFW
                
        'coloca os dados na tela
        GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_Regra_Col) = objRegraWFW.sRegra
        
        If iTransacao = TRANSACAOWFW_PEDIDO_COMPRA Or iTransacao = TRANSACAOWFW_PEDIDO_VENDA Or iTransacao = TRANSACAOWFW_ORCAMENTO_VENDA Or iTransacao = TRANSACAOWFW_ORCAMENTO_SERVICO Then
        
            'Carrega a combo de Tipo de Bloqueio
            For iIndice = 0 To TipoBloqueio.ListCount - 1
        
                If TipoBloqueio.ItemData(iIndice) = objRegraWFW.iTipoBloqueio Then
                    GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_TipoBloqueio_Col) = TipoBloqueio.List(iIndice)
                    Exit For
                End If
                
            Next
        
        End If
        
        If objRegraWFW.sUsuario = "" Then
            GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_ValidoPara_Col) = ValidoPara.List(1)
        Else
            GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_ValidoPara_Col) = ValidoPara.List(0)
        End If
        
        If Len(objRegraWFW.sEmailPara) > 0 Or Len(objRegraWFW.sEmailAssunto) > 0 Or Len(objRegraWFW.sEmailMsg) > 0 Then
            GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_Email_Col) = MARCADO
        End If
        
        If Len(objRegraWFW.sAvisoMsg) > 0 Or objRegraWFW.colUsuarios.Count > 0 Then
            GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_Aviso_Col) = MARCADO
        End If
        
        If Len(objRegraWFW.sLogDoc) > 0 Or Len(objRegraWFW.sLogMsg) > 0 Then
            GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_Log_Col) = MARCADO
        End If
        
        If Len(Trim(objRegraWFW.sRelModulo)) > 0 Or Len(Trim(objRegraWFW.sRelNome)) > 0 Or Len(Trim(objRegraWFW.sRelOpcao)) > 0 Then
            GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_Rel_Col) = MARCADO
        End If
        
        If Len(Trim(objRegraWFW.sBrowseModulo)) > 0 Or Len(Trim(objRegraWFW.sBrowseNome)) > 0 Or Len(Trim(objRegraWFW.sBrowseOpcao)) > 0 Then
            GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_Browse_Col) = MARCADO
        End If
        
        GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_EmailParaGrid_Col) = objRegraWFW.sEmailPara
        GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_EmailAssuntoGrid_Col) = objRegraWFW.sEmailAssunto
        GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_EmailMsgGrid_Col) = objRegraWFW.sEmailMsg
        GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_AvisoMsgGrid_Col) = objRegraWFW.sAvisoMsg
        GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_LogDocGrid_Col) = objRegraWFW.sLogDoc
        GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_LogMsgGrid_Col) = objRegraWFW.sLogMsg
        If Len(Trim(objRegraWFW.sRelModulo)) > 0 Then
            GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_RelModulo_Col) = gcolModulo.Nome(objRegraWFW.sRelModulo)
        End If
        GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_RelNome_Col) = objRegraWFW.sRelNome
        GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_RelOpcao_Col) = objRegraWFW.sRelOpcao
        GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_RelPorEmail_Col) = CStr(objRegraWFW.iRelPorEmail)
        GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_RelSel_Col) = objRegraWFW.sRelSel
        GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_RelAnexo_Col) = objRegraWFW.sRelAnexo
        
        If Len(Trim(objRegraWFW.sBrowseModulo)) > 0 Then
            GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_BrowseModulo_Col) = gcolModulo.Nome(objRegraWFW.sBrowseModulo)
        End If
        
        GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_BrowseNome_Col) = objRegraWFW.sBrowseNome
        GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_BrowseOpcao_Col) = objRegraWFW.sBrowseOpcao
        

        
        
        'leitura dos usuarios do item em questão
        lErro = CF("AvisoUsuWFW_Le_Item", objRegraWFW)
        If lErro <> SUCESSO Then gError 178087
        
        For iIndice1 = 1 To objRegraWFW.colUsuarios.Count
            sUsuario = objRegraWFW.colUsuarios(iIndice1)
        
            GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_AvisoUsuGrid_Col) = GridRegras.TextMatrix(objRegraWFW.iItem, iGrid_AvisoUsuGrid_Col) & sUsuario & " "
        
        Next
        
        objGridRegra.iLinhasExistentes = objGridRegra.iLinhasExistentes + 1
            
    Next
    
    giGridRefresh = 1
    Call Grid_Refresh_Checkbox(objGridRegra)
    giGridRefresh = 0
    
    Call Limpa_Tela(Me)
    ComboModulo.ListIndex = -1
    CodRelatorio.Text = ""
    ComboOpcoes.Text = ""
    ComboModuloBrowser.ListIndex = -1
    CodBrowser.Text = ""
    ComboOpcaoBrowser.Text = ""
    
    Carga_Grid = SUCESSO

    Exit Function

Erro_Carga_Grid:

    Carga_Grid = gErr

    Select Case gErr

        Case 178072 To 178074, 178087
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178080)

    End Select
    
    Exit Function

End Function

Private Sub Posiciona_Texto_Tela(objControl As Control, sTexto As String)
'posiciona o texto sTexto no controle objControl da tela

Dim iPos As Integer
Dim iTamanho As Integer
Dim objGrid As Object

    iPos = objControl.SelStart
    objControl.Text = Mid(objControl.Text, 1, iPos) & sTexto & Mid(objControl.Text, iPos + 1, Len(objControl.Text))
    objControl.SelStart = iPos + Len(sTexto)
    
    If Not (Me.ActiveControl Is objControl) Then
    
        If iFrameAtual = TAB_REGRAS Then
            Set objGrid = GridRegras
        
            If iPos >= Len(objGrid.TextMatrix(objGrid.Row, objGrid.Col)) Then
                iTamanho = 0
            Else
                iTamanho = Len(objGrid.TextMatrix(objGrid.Row, objGrid.Col)) - iPos
            End If
            objGrid.TextMatrix(objGrid.Row, objGrid.Col) = Mid(objGrid.TextMatrix(objGrid.Row, objGrid.Col), 1, iPos) & sTexto & Mid(objGrid.TextMatrix(objGrid.Row, objGrid.Col), iPos + 1, iTamanho)
        
        End If
        
    End If

    iAlterado = REGISTRO_ALTERADO

End Sub

Function Funcoes1(sFuncao As String) As Long

Dim iPos As Integer

On Error GoTo Erro_Funcoes1

    If iFrameAtual = TAB_REGRAS Then
    
        If GridRegras.Row > 0 And GridRegras.Row <= objGridRegra.iLinhasExistentes + 1 And GridRegras.Col > 0 Then
            
            Select Case GridRegras.Col
            
                Case iGrid_Regra_Col
                    Call Posiciona_Texto_Tela(Regra, Funcoes.Text)
                            
            End Select
            
        End If
        
    ElseIf iFrameAtual = TAB_EMAIL Then
    
        If objCampoAtual Is EmailPara Or objCampoAtual Is EmailAssunto Or objCampoAtual Is EmailMsg Then
            Call Posiciona_Texto_Tela(objCampoAtual, Funcoes.Text)
        End If
    
    ElseIf iFrameAtual = TAB_AVISO Then
        If objCampoAtual Is AvisoMsg Then
            Call Posiciona_Texto_Tela(objCampoAtual, Funcoes.Text)
        End If
    ElseIf iFrameAtual = TAB_LOG Then
        If objCampoAtual Is LogDoc Or objCampoAtual Is LogMsg Then
            Call Posiciona_Texto_Tela(objCampoAtual, Funcoes.Text)
        End If
    ElseIf iFrameAtual = TAB_RELATORIO Then
        If objCampoAtual Is RelSel Or objCampoAtual Is RelAnexo Then
            Call Posiciona_Texto_Tela(objCampoAtual, Funcoes.Text)
        End If
    End If
    
    Funcoes1 = SUCESSO
    
    Exit Function
    
Erro_Funcoes1:

    Funcoes1 = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178104)
            
    End Select
        
    Exit Function

End Function

Function Mnemonicos1(sMnemonico As String) As Long

Dim iPos As Integer

On Error GoTo Erro_Mnemonicos1

    If iFrameAtual = TAB_REGRAS Then
    
        If GridRegras.Row > 0 And GridRegras.Row <= objGridRegra.iLinhasExistentes + 1 And GridRegras.Col > 0 Then
            
            Select Case GridRegras.Col
            
                Case iGrid_Regra_Col
                    Call Posiciona_Texto_Tela(Regra, Mnemonicos.Text)
                    If GridRegras.Row - GridRegras.FixedRows = objGridRegra.iLinhasExistentes Then
                        objGridRegra.iLinhasExistentes = objGridRegra.iLinhasExistentes + 1
                    End If
                    
                    
            End Select
            
        End If
        
    ElseIf iFrameAtual = TAB_EMAIL Then
        
        If objCampoAtual Is EmailPara Or objCampoAtual Is EmailAssunto Or objCampoAtual Is EmailMsg Then
            Call Posiciona_Texto_Tela(objCampoAtual, Mnemonicos.Text)
        End If
    
    ElseIf iFrameAtual = TAB_AVISO Then
        If objCampoAtual Is AvisoMsg Then
            Call Posiciona_Texto_Tela(objCampoAtual, Mnemonicos.Text)
        End If
    ElseIf iFrameAtual = TAB_LOG Then
        If objCampoAtual Is LogDoc Or objCampoAtual Is LogMsg Then
            Call Posiciona_Texto_Tela(objCampoAtual, Mnemonicos.Text)
        End If
    ElseIf iFrameAtual = TAB_RELATORIO Then
        If objCampoAtual Is RelSel Or objCampoAtual Is RelAnexo Then
            Call Posiciona_Texto_Tela(objCampoAtual, Mnemonicos.Text)
        End If
    End If

    Mnemonicos1 = SUCESSO
    
    Exit Function
    
Erro_Mnemonicos1:

    Mnemonicos1 = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178105)
            
    End Select
        
    Exit Function

End Function

Function Operadores1() As Long

Dim iPos As Integer

On Error GoTo Erro_Operadores1

    If iFrameAtual = TAB_REGRAS Then
    
        If GridRegras.Row > 0 And GridRegras.Row <= objGridRegra.iLinhasExistentes + 1 And GridRegras.Col > 0 Then
            
            Select Case GridRegras.Col
            
                Case iGrid_Regra_Col
                    Call Posiciona_Texto_Tela(Regra, Operadores.Text)
                            
            End Select
            
        End If
        
    ElseIf iFrameAtual = TAB_EMAIL Then
    
        If objCampoAtual Is EmailPara Or objCampoAtual Is EmailAssunto Or objCampoAtual Is EmailMsg Then
            Call Posiciona_Texto_Tela(objCampoAtual, Operadores.Text)
        End If
    
    ElseIf iFrameAtual = TAB_AVISO Then
        If objCampoAtual Is AvisoMsg Then
            Call Posiciona_Texto_Tela(objCampoAtual, Operadores.Text)
        End If
    ElseIf iFrameAtual = TAB_LOG Then
        If objCampoAtual Is LogDoc Or objCampoAtual Is LogMsg Then
            Call Posiciona_Texto_Tela(objCampoAtual, Operadores.Text)
        End If
    ElseIf iFrameAtual = TAB_RELATORIO Then
        If objCampoAtual Is RelSel Or objCampoAtual Is RelAnexo Then
            Call Posiciona_Texto_Tela(objCampoAtual, Operadores.Text)
        End If
    End If
    
    Operadores1 = SUCESSO
    
    Exit Function
    
Erro_Operadores1:

    Operadores1 = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178106)
            
    End Select
        
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Workflow"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Workflow"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Private Sub Unload(objme As Object)
   ' Parent.UnloadDoFilho
    
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

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
        
    
    ElseIf KeyCode = KEYCODE_VERIFICAR_SINTAXE Then
    
        If Checkbox_Verifica_Sintaxe.Value = MARCADO Then
            Checkbox_Verifica_Sintaxe.Value = DESMARCADO
        Else
            Checkbox_Verifica_Sintaxe.Value = MARCADO
        End If
            
    End If
    

End Sub

Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

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

Private Sub ComboModulo_Click()

Dim lErro As Long
Dim colRelatorio As New Collection
Dim sCodRel As String
Dim vntCodRel As Variant
Dim sModulo As String

On Error GoTo Erro_ComboModulo_Click

    If ComboModulo.ListIndex <> -1 Then
        
        sModulo = ComboModulo.Text
        
        'Preenche a colecao com os nomes dos relatorios existentes no BD
        lErro = CF("Relatorios_Le_NomeModulo", sModulo, colRelatorio)
        If lErro <> SUCESSO Then gError 178379

        CodRelatorio.Clear
    
        For Each vntCodRel In colRelatorio
            sCodRel = vntCodRel
            CodRelatorio.AddItem (sCodRel)
        Next
        
        ComboOpcoes.Clear

    End If
    
    Exit Sub
     
Erro_ComboModulo_Click:

    Select Case gErr
          
        Case 178379
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 178380)
     
    End Select
     
    Exit Sub

End Sub

Private Sub ComboModuloBrowser_Click()

Dim lErro As Long
Dim colTela As New Collection
Dim vTela As Variant
Dim sModulo As String

On Error GoTo Erro_ComboModuloBrowser_Click
    
    If ComboModuloBrowser.ListIndex <> -1 Then
    
        sModulo = ComboModuloBrowser.Text
    
        CodBrowser.Clear
    
        'Lê os nomes de telas de Browse na tabela BrowseArquivo
        lErro = CF("BrowseArquivo_Le_Telas", sModulo, colTela)
        If lErro <> SUCESSO Then gError 178382
        
        'Preenche a list da ComboBox Tela com os nomes lidos
        For Each vTela In colTela
            CodBrowser.AddItem vTela
        Next

        ComboOpcaoBrowser.Clear

    End If
    
    Exit Sub

Erro_ComboModuloBrowser_Click:

    Select Case gErr

        Case 178382

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178383)

    End Select

    Exit Sub

End Sub

Public Sub ValidoPara_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValidoPara_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRegra)
    
End Sub

Public Sub ValidoPara_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegra)
    
End Sub

Public Sub ValidoPara_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegra.objControle = ValidoPara
    lErro = Grid_Campo_Libera_Foco(objGridRegra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

