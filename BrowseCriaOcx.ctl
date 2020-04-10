VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.UserControl BrowseCria 
   ClientHeight    =   7050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11835
   LockControls    =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   11835
   Begin VB.CommandButton BotaoValidarCodigo 
      Caption         =   "Validar Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9840
      TabIndex        =   95
      Top             =   1290
      Width           =   900
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   405
      Left            =   10800
      TabIndex        =   94
      Top             =   690
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.TextBox NomeTelaConsulta 
      Height          =   285
      Left            =   5355
      TabIndex        =   12
      ToolTipText     =   "Digite o nome da tela de cadastro da tabela."
      Top             =   1695
      Width           =   1815
   End
   Begin VB.Frame Frame5 
      Caption         =   "O Browse tem os botões:"
      Height          =   600
      Left            =   135
      TabIndex        =   91
      Top             =   1380
      Width           =   3510
      Begin VB.CheckBox Botao 
         Caption         =   "Consultar"
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
         Index           =   2
         Left            =   2265
         TabIndex        =   10
         Top             =   225
         Width           =   1200
      End
      Begin VB.CheckBox Botao 
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
         Height          =   255
         Index           =   1
         Left            =   1425
         TabIndex        =   9
         Top             =   225
         Value           =   1  'Checked
         Width           =   960
      End
      Begin VB.CheckBox Botao 
         Caption         =   "Selecionar"
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
         Left            =   150
         TabIndex        =   8
         Top             =   210
         Value           =   1  'Checked
         Width           =   1410
      End
   End
   Begin VB.CommandButton BotaoAcertarData 
      Caption         =   "Acertar Default Data Nula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Limpar"
      Top             =   1125
      Width           =   975
   End
   Begin VB.Frame Frame4 
      Caption         =   "Validar Browses"
      Height          =   1215
      Left            =   8460
      TabIndex        =   90
      Top             =   15
      Width           =   2265
      Begin VB.CommandButton BotaoTeste 
         Caption         =   "Testar"
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
         Left            =   1170
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Abre o browse informado"
         Top             =   855
         Width           =   1080
      End
      Begin VB.CheckBox optCorrecao 
         Caption         =   "Script Correção"
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
         Left            =   30
         TabIndex        =   49
         Top             =   630
         Width           =   2100
      End
      Begin VB.CommandButton BotaoVerificar 
         Caption         =   "Validar"
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
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Valida os Browses Cadastrados no sistema"
         Top             =   855
         Width           =   1080
      End
      Begin VB.CheckBox optApenasBrowse 
         Caption         =   "Apenas Browse Tela"
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
         Left            =   30
         TabIndex        =   48
         Top             =   420
         Width           =   2100
      End
      Begin VB.CheckBox optAviso 
         Caption         =   "Exibir Avisos"
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
         Left            =   30
         TabIndex        =   47
         Top             =   225
         Width           =   1650
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Modulos"
      Height          =   1215
      Left            =   5325
      TabIndex        =   76
      Top             =   15
      Width           =   3120
      Begin VB.ComboBox ModuloTela 
         Height          =   315
         ItemData        =   "BrowseCriaOcx.ctx":0000
         Left            =   2325
         List            =   "BrowseCriaOcx.ctx":0022
         Sorted          =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "Digite o Módulo(VB) da Tela de Cadastro da Tabela."
         Top             =   765
         Width           =   735
      End
      Begin VB.ComboBox ModuloAcesso 
         Height          =   315
         ItemData        =   "BrowseCriaOcx.ctx":0059
         Left            =   2325
         List            =   "BrowseCriaOcx.ctx":007B
         Sorted          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "Digite o Módulo a que o Browse Vai Pertencer."
         Top             =   315
         Width           =   735
      End
      Begin VB.ComboBox ModuloFormata 
         Height          =   315
         ItemData        =   "BrowseCriaOcx.ctx":00B2
         Left            =   825
         List            =   "BrowseCriaOcx.ctx":00DD
         Sorted          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "Digite ou Escolha o Módulo(VB) que está a Classe de Formatação dessa Tabela."
         Top             =   750
         Width           =   735
      End
      Begin VB.ComboBox ModuloClasse 
         Height          =   315
         ItemData        =   "BrowseCriaOcx.ctx":0124
         Left            =   825
         List            =   "BrowseCriaOcx.ctx":014F
         Sorted          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "Digite o Mólulo(VB) Onde se Encontra a Classe da Tabela."
         Top             =   315
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Tela:"
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
         Left            =   1845
         TabIndex        =   80
         Top             =   795
         Width           =   495
      End
      Begin VB.Label Label13 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   79
         Top             =   345
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Formata:"
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
         Left            =   30
         TabIndex        =   78
         Top             =   795
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Acesso:"
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
         Left            =   1605
         TabIndex        =   77
         Top             =   345
         Width           =   735
      End
   End
   Begin VB.TextBox UltErro 
      Height          =   285
      Left            =   8280
      TabIndex        =   13
      ToolTipText     =   "Digite qual foi o Ultimo Erro utilizado por você."
      Top             =   1290
      Width           =   1530
   End
   Begin VB.TextBox DescBrowse 
      Height          =   285
      Left            =   1065
      TabIndex        =   3
      ToolTipText     =   "Digite a Descrição da Tabela"
      Top             =   1005
      Width           =   4200
   End
   Begin VB.TextBox NomeTela 
      Height          =   285
      Left            =   5355
      TabIndex        =   11
      ToolTipText     =   "Digite o nome da tela de cadastro da tabela."
      Top             =   1290
      Width           =   1815
   End
   Begin VB.TextBox NomeBrowse 
      Height          =   285
      Left            =   3720
      TabIndex        =   2
      ToolTipText     =   "Digite o Nome do Browse que Vai Ser Criado."
      Top             =   525
      Width           =   1560
   End
   Begin VB.TextBox Classe 
      Height          =   315
      Left            =   1035
      TabIndex        =   1
      ToolTipText     =   "Digite o Nome da Classe da Relacionada a View ou a Tabela."
      Top             =   540
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Código"
      Height          =   1455
      Left            =   120
      TabIndex        =   59
      Top             =   5535
      Width           =   11625
      Begin VB.TextBox ScriptDic 
         Height          =   690
         Left            =   10485
         MultiLine       =   -1  'True
         TabIndex        =   45
         ToolTipText     =   "Script de Criação do Browse"
         Top             =   300
         Width           =   1050
      End
      Begin VB.CommandButton BotaoExpDic 
         Caption         =   "Exportar"
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
         Left            =   10485
         TabIndex        =   46
         ToolTipText     =   "Script de Rotinas de Leitura, Exclusão e Gravação."
         Top             =   1020
         Width           =   1050
      End
      Begin VB.TextBox ScriptExclusao 
         Height          =   690
         Left            =   9420
         MultiLine       =   -1  'True
         TabIndex        =   43
         ToolTipText     =   "Script de Criação do Browse"
         Top             =   300
         Width           =   1050
      End
      Begin VB.CommandButton BotaoExpExclusao 
         Caption         =   "Exportar"
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
         Left            =   9420
         TabIndex        =   44
         ToolTipText     =   "Script de Rotinas de Leitura, Exclusão e Gravação."
         Top             =   1020
         Width           =   1050
      End
      Begin VB.TextBox ScriptGravacao 
         Height          =   690
         Left            =   8340
         MultiLine       =   -1  'True
         TabIndex        =   41
         ToolTipText     =   "Script de Criação do Browse"
         Top             =   300
         Width           =   1050
      End
      Begin VB.CommandButton BotaoExpGravacao 
         Caption         =   "Exportar"
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
         Left            =   8340
         TabIndex        =   42
         ToolTipText     =   "Script de Rotinas de Leitura, Exclusão e Gravação."
         Top             =   1020
         Width           =   1050
      End
      Begin VB.TextBox ScriptLeitura 
         Height          =   690
         Left            =   7275
         MultiLine       =   -1  'True
         TabIndex        =   39
         ToolTipText     =   "Script de Criação do Browse"
         Top             =   300
         Width           =   1050
      End
      Begin VB.CommandButton BotaoExpLeitura 
         Caption         =   "Exportar"
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
         Left            =   7275
         TabIndex        =   40
         ToolTipText     =   "Script de Rotinas de Leitura, Exclusão e Gravação."
         Top             =   1020
         Width           =   1050
      End
      Begin VB.TextBox ScriptType 
         Height          =   690
         Left            =   6210
         MultiLine       =   -1  'True
         TabIndex        =   37
         ToolTipText     =   "Script de Criação do Browse"
         Top             =   300
         Width           =   1050
      End
      Begin VB.CommandButton BotaoExpType 
         Caption         =   "Exportar"
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
         Left            =   6210
         TabIndex        =   38
         ToolTipText     =   "Script de Rotinas de Leitura, Exclusão e Gravação."
         Top             =   1020
         Width           =   1050
      End
      Begin VB.Frame Frame2 
         Caption         =   "Gerar"
         Height          =   1140
         Left            =   180
         TabIndex        =   83
         Top             =   210
         Width           =   4890
         Begin VB.CommandButton BotaoIncluirGrid 
            Caption         =   "Grid"
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
            Left            =   3825
            TabIndex        =   32
            ToolTipText     =   "Insere Grid"
            Top             =   150
            Width           =   990
         End
         Begin VB.CommandButton BotaoIncluirTab 
            Caption         =   "Tab"
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
            Left            =   3825
            TabIndex        =   33
            ToolTipText     =   "Insere Tab"
            Top             =   465
            Width           =   990
         End
         Begin VB.CheckBox optTodos 
            Caption         =   "TODOS"
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
            Left            =   2550
            TabIndex        =   31
            Top             =   810
            Width           =   1050
         End
         Begin VB.CommandButton BotaoGerarCodigo 
            Caption         =   "Gerar"
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
            Left            =   3825
            TabIndex        =   34
            ToolTipText     =   "Gera Script de Criação do Browse"
            Top             =   795
            Width           =   990
         End
         Begin VB.CheckBox optTela 
            Caption         =   "Tela"
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
            Left            =   2550
            TabIndex        =   30
            Top             =   540
            Width           =   1050
         End
         Begin VB.CheckBox optDic 
            Caption         =   "Dic"
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
            Left            =   2550
            TabIndex        =   29
            Top             =   270
            Width           =   1050
         End
         Begin VB.CheckBox optExclusao 
            Caption         =   "Exclusão"
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
            Left            =   1245
            TabIndex        =   28
            Top             =   810
            Width           =   1230
         End
         Begin VB.CheckBox optGravacao 
            Caption         =   "Gravação"
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
            Left            =   1245
            TabIndex        =   27
            Top             =   540
            Width           =   1275
         End
         Begin VB.CheckBox optLeitura 
            Caption         =   "Leitura"
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
            Left            =   1245
            TabIndex        =   26
            Top             =   270
            Width           =   1050
         End
         Begin VB.CheckBox optClasse 
            Caption         =   "Classe"
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
            Left            =   195
            TabIndex        =   25
            Top             =   810
            Width           =   1050
         End
         Begin VB.CheckBox optType 
            Caption         =   "Type"
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
            Left            =   195
            TabIndex        =   24
            Top             =   540
            Width           =   1050
         End
         Begin VB.CheckBox optBrowse 
            Caption         =   "Browse"
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
            Left            =   195
            TabIndex        =   23
            Top             =   270
            Width           =   1050
         End
      End
      Begin VB.CommandButton botaoExpBrowse 
         Caption         =   "Exportar"
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
         Left            =   5145
         TabIndex        =   36
         ToolTipText     =   "Script de Rotinas de Leitura, Exclusão e Gravação."
         Top             =   1020
         Width           =   1050
      End
      Begin VB.TextBox ScriptBrowse 
         Height          =   690
         Left            =   5145
         MultiLine       =   -1  'True
         TabIndex        =   35
         ToolTipText     =   "Script de Criação do Browse"
         Top             =   300
         Width           =   1050
      End
      Begin VB.Label Label17 
         Caption         =   "Dic:"
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
         Left            =   10485
         TabIndex        =   89
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label16 
         Caption         =   "Exclusão:"
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
         Left            =   9420
         TabIndex        =   88
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label10 
         Caption         =   "Gravação:"
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
         Left            =   8340
         TabIndex        =   87
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label9 
         Caption         =   "Leitura:"
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
         Left            =   7275
         TabIndex        =   86
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label8 
         Caption         =   "Type:"
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
         Left            =   6210
         TabIndex        =   85
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label7 
         Caption         =   "Browse:"
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
         Left            =   5145
         TabIndex        =   84
         Top             =   120
         Width           =   1035
      End
   End
   Begin VB.ComboBox NomeArq 
      Height          =   315
      ItemData        =   "BrowseCriaOcx.ctx":0196
      Left            =   1050
      List            =   "BrowseCriaOcx.ctx":0198
      TabIndex        =   0
      ToolTipText     =   "Digite ou Escolha o Nome da View ou Tabela do Banco de Dados."
      Top             =   75
      Width           =   4245
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   10800
      ScaleHeight     =   495
      ScaleWidth      =   915
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   105
      Width           =   975
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   45
         Picture         =   "BrowseCriaOcx.ctx":019A
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   495
         Picture         =   "BrowseCriaOcx.ctx":06CC
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Campos"
      Height          =   3480
      Left            =   120
      TabIndex        =   57
      Top             =   2040
      Width           =   11625
      Begin MSMask.MaskEdBox SubTipo 
         Height          =   225
         Left            =   6570
         TabIndex        =   93
         ToolTipText     =   "Altere, se For Necessário, a Precisão do Campo."
         Top             =   1395
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.CommandButton MarcaChave 
         Caption         =   "Marcar Chaves"
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
         Left            =   6120
         TabIndex        =   19
         ToolTipText     =   "Marca Todos Campos Para Serem Chaves."
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton DesmarcaChave 
         Caption         =   "Desmarcar Chaves"
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
         Left            =   7320
         TabIndex        =   20
         ToolTipText     =   "Desmarca Todos Campos Para Serem Chaves."
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton DesmarcarIndices 
         Caption         =   "Desmarcar Indices"
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
         Left            =   4320
         TabIndex        =   18
         ToolTipText     =   "Desmarca Todos Campos Para Serem Índices."
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton MarcarIndices 
         Caption         =   "Marcar Indices"
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
         Left            =   3120
         TabIndex        =   17
         ToolTipText     =   "Marca Todos Campos Para Serem Índices."
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton DesmarcarBrowse 
         Caption         =   "Desmarcar Browses"
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
         Left            =   1260
         TabIndex        =   16
         ToolTipText     =   "Desmarca Todos Campos Para Serem Exibidos no Browse"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton MarcarBrowse 
         Caption         =   "Marcar Browses"
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
         Left            =   75
         TabIndex        =   15
         ToolTipText     =   "Marca Todos Campos Para Serem Exibidos no Browse."
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton DesmarcarClasses 
         Caption         =   "Desmarcar Classes"
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
         Left            =   10470
         TabIndex        =   22
         ToolTipText     =   "Desmarca Todos Campos Como Fazendo Parte da Classe"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton MarcarClasse 
         Caption         =   "Marcar Classes"
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
         Left            =   9255
         TabIndex        =   21
         ToolTipText     =   "Marca Todos Campos Como Fazendo Parte da Classe."
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CheckBox Chave 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   4440
         TabIndex        =   74
         ToolTipText     =   "Marque as Chaves Únicas que Serão Utilizadas na Leitura da Tabela."
         Top             =   840
         Width           =   690
      End
      Begin MSMask.MaskEdBox Descricao 
         Height          =   225
         Left            =   4320
         TabIndex        =   70
         ToolTipText     =   "Altere, se For Necessário, a Descrição da Coluna. Isso Vai aparecer no Título de Cada Campo."
         Top             =   1560
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.CheckBox Indice 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   4440
         TabIndex        =   72
         ToolTipText     =   "Marque se Vai Ser Criado um Índice Para Esse Campo no Browse."
         Top             =   600
         Width           =   690
      End
      Begin VB.CheckBox TemClasse 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   4440
         TabIndex        =   69
         ToolTipText     =   "Marque se for um atributo da Classe Digitada."
         Top             =   360
         Width           =   690
      End
      Begin MSMask.MaskEdBox TamanhoTela 
         Height          =   225
         Left            =   4440
         TabIndex        =   67
         ToolTipText     =   "Altere, se For Necessário, o Tamanho Ocupado Pelo Campo no Browse."
         Top             =   2040
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.CheckBox Browse 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   4440
         TabIndex        =   65
         ToolTipText     =   "Marque se For Para Aparecer no Browse."
         Top             =   1080
         Width           =   690
      End
      Begin MSMask.MaskEdBox AtribClasse 
         Height          =   225
         Left            =   4350
         TabIndex        =   64
         ToolTipText     =   "Altere, se For Necessário, o Nome do Atributo da Classe."
         Top             =   1320
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Ordem 
         Height          =   225
         Left            =   3840
         TabIndex        =   63
         ToolTipText     =   "Altere, se For Necessário, a Ordem do Campo."
         Top             =   2160
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Precisao 
         Height          =   225
         Left            =   3105
         TabIndex        =   62
         ToolTipText     =   "Altere, se For Necessário, a Precisão do Campo."
         Top             =   2280
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Tamanho 
         Height          =   225
         Left            =   4200
         TabIndex        =   61
         ToolTipText     =   "Altere, se For Necessário, o Tamanho do Campo."
         Top             =   2400
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Tipo 
         Height          =   225
         Left            =   4440
         TabIndex        =   60
         ToolTipText     =   "Tipo de Dado Utilizado pelo SQL Server."
         Top             =   1800
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Coluna 
         Height          =   225
         Left            =   1200
         TabIndex        =   55
         ToolTipText     =   "Nome da Coluna no Banco de Dados"
         Top             =   2280
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridColunas 
         Height          =   2580
         Left            =   60
         TabIndex        =   14
         Top             =   270
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   4551
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
   End
   Begin VB.Label Label18 
      Caption         =   "Tela de Consulta:"
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
      Left            =   3810
      TabIndex        =   92
      Top             =   1725
      Width           =   1590
   End
   Begin VB.Label ErroGerado 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   8295
      TabIndex        =   82
      ToolTipText     =   "Ultimo Erro Gerado pelo Sistema."
      Top             =   1680
      Width           =   1515
   End
   Begin VB.Label Label15 
      Caption         =   "Gerou Até:"
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
      Left            =   7275
      TabIndex        =   81
      Top             =   1725
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Último Erro:"
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
      Left            =   7200
      TabIndex        =   75
      Top             =   1335
      Width           =   1095
   End
   Begin VB.Label Label6 
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
      Height          =   255
      Left            =   120
      TabIndex        =   73
      Top             =   1065
      Width           =   990
   End
   Begin VB.Label Label2 
      Caption         =   "Tela de Edição:"
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
      Left            =   3960
      TabIndex        =   71
      Top             =   1350
      Width           =   1395
   End
   Begin VB.Label Label5 
      Caption         =   "Browse:"
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
      Left            =   3030
      TabIndex        =   68
      Top             =   585
      Width           =   735
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   420
      TabIndex        =   66
      Top             =   570
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Arquivo:"
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
      Left            =   300
      TabIndex        =   58
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "BrowseCria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Declaração de variáveis globais
Dim iAlterado As Integer
Dim iAlteradoNomeArq As Integer
Dim iCont As Integer
Dim iContAdicao As Integer
Dim gsErroLeitura As String
Dim gsNomeArqAnterior As String
Dim gbTemTab As Boolean
Dim gbTemGrid As Boolean

Dim gobjTela As ClassCriaTela

Dim objGridColunas As AdmGrid
Dim iGrid_Browse_Col As Integer
Dim iGrid_Coluna_Col As Integer
Dim iGrid_Tipo_Col As Integer
Dim iGrid_Tamanho_Col As Integer
Dim iGrid_TamanhoTela_Col As Integer
Dim iGrid_Precisao_Col As Integer
Dim iGrid_Ordem_Col As Integer
Dim iGrid_AtribClasse_Col As Integer
Dim iGrid_TemClasse_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_Indice_Col As Integer
Dim iGrid_Chave_Col As Integer
Dim iGrid_SubTipo_Col As Integer

Dim gsTipoArquivo As String
Dim gbCriarScriptDelete As Boolean

Type typeColunasTabelas
     sArquivo As String
     sArquivoTipo As String
     sColuna As String
     sColunaTipo As String
     lColunaTamanho As Long
     lColunaPrecisao As Long
End Type

Const ARQUIVO_TABELA = "U"
Const ARQUIVO_VIEW = "V"
Const TECLA_TAB = "    "
Const STRING_STRING_MAX = 255

Const TIPO_GRID = 1
Const TIPO_FRAME = 2
Const TIPO_OUTRO = 3

Private WithEvents objEvento As AdmEvento
Attribute objEvento.VB_VarHelpID = -1

Private Function GerarBrowse(ByVal colColunasTabelas As Collection) As Long

Dim lErro As Long
Dim sBrowse As String
Dim sNL As String

On Error GoTo Erro_GerarBrowse

    sNL = Chr(10)

    lErro = Browse_Cria(colColunasTabelas, sBrowse)
    If lErro <> SUCESSO Then gError 131711
    
    If gbCriarScriptDelete Then
        Call Browse_Apaga
        ScriptBrowse.Text = ScriptBrowse.Text & sNL & sNL & sBrowse
    Else
        ScriptBrowse.Text = sBrowse
    End If
    
    GerarBrowse = SUCESSO

    Exit Function

Erro_GerarBrowse:

    GerarBrowse = gErr

    Select Case gErr
    
        Case 131711

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143899)

    End Select
    
    Exit Function
    
End Function

Private Function GerarType(ByVal colColunasTabelas As Collection) As Long

Dim lErro As Long
Dim sClasse As String
Dim sType As String
Dim sNL As String

On Error GoTo Erro_GerarType

    lErro = ClasseType_Cria(colColunasTabelas, sClasse, sType)
    If lErro <> SUCESSO Then gError 131714
    
    ScriptType.Text = sType
    
    GerarType = SUCESSO
    
    Exit Function

Erro_GerarType:

    GerarType = gErr

    Select Case gErr
    
        Case 131714

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143900)

    End Select
    
    Exit Function
    
End Function

Private Function GerarClasse(ByVal colColunasTabelas As Collection) As Long

Dim lErro As Long
Dim sClasse As String
Dim sType As String
Dim sNL As String

On Error GoTo Erro_GerarClasse

    lErro = ClasseType_Cria(colColunasTabelas, sClasse, sType)
    If lErro <> SUCESSO Then gError 131714
    
    Open CurDir & "\" & Classe.Text & ".cls" For Output As #3
    
    Print #3, "VERSION 1.0 CLASS"
    Print #3, "BEGIN"
    Print #3, "  MultiUse = -1  'True"
    Print #3, "  Persistable = 0  'NotPersistable"
    Print #3, "  DataBindingBehavior = 0  'vbNone"
    Print #3, "  DataSourceBehavior = 0   'vbNone"
    Print #3, "  MTSTransactionMode = 0   'NotAnMTSObject"
    Print #3, "End"
    Print #3, "Attribute VB_Name = " & """" & Classe.Text & """"
    Print #3, "Attribute VB_GlobalNameSpace = False"
    Print #3, "Attribute VB_Creatable = True"
    Print #3, "Attribute VB_PredeclaredId = False"
    Print #3, "Attribute VB_Exposed = True"
    Print #3, "Attribute VB_Ext_KEY = " & """" & "SavedWithClassBuilder6" & """" & " ," & """" & "Yes" & """"
    Print #3, "Attribute VB_Ext_KEY = " & """" & "Top_Level" & """" & ", " & """" & "; Yes; " & """"""
    Print #3, "Option Explicit"
    Print #3, ""
    Print #3, sClasse
    
    Close #3
    
    GerarClasse = SUCESSO
    
    Exit Function

Erro_GerarClasse:

    GerarClasse = gErr

    Select Case gErr
    
        Case 131714

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143901)

    End Select
    
    Close #3
    
    Exit Function
    
End Function

Private Sub CalculaProximoErro(sProximoErro As String)
'Verifica qual foi o último erro preenchido na tela e calcula os próximos
'utilizados nas funções => Se não estiver preenchido deixa indicado no script

    iCont = iCont + 1
    
    If Len(Trim(UltErro.Text)) <> 0 Then
    
        sProximoErro = CStr(StrParaLong(UltErro.Text) + iCont)
    
    Else
        sProximoErro = "ULTIMO_ERRO + " & CStr(iCont)
    
    End If

End Sub

Private Function UltimoErro() As String
'Obtém o último erro usado

    If Len(Trim(UltErro.Text)) <> 0 Then
    
        UltimoErro = CStr(StrParaLong(UltErro.Text) + iCont)
    
    Else
        UltimoErro = "ULTIMO_ERRO + " & CStr(iCont)
    
    End If

End Function

Private Sub Marca_Desmarca(ByVal iColuna As Integer, ByVal iValor As Integer)
'Marca ou desmarca uma coluna do Grid Passada por parametro

Dim iIndice As Integer
Dim sSigla As String
Dim sTipo As String
Dim sNome As String
Dim iTipo As Integer
Dim sTipoVB As String

    For iIndice = 1 To objGridColunas.iLinhasExistentes
    
        GridColunas.TextMatrix(iIndice, iColuna) = iValor
        
        If iColuna = iGrid_TemClasse_Col Then
        
            If iValor = MARCADO Then
            
                sTipo = GridColunas.TextMatrix(iIndice, iGrid_Tipo_Col)
                sNome = GridColunas.TextMatrix(iIndice, iGrid_Coluna_Col)
            
                Call ObtemSiglaTipo(sTipo, sSigla, iTipo, sTipoVB)
                GridColunas.TextMatrix(iIndice, iGrid_AtribClasse_Col) = sSigla & sNome
            
            Else
                GridColunas.TextMatrix(iIndice, iGrid_AtribClasse_Col) = ""
            
            End If
        
        End If
        
    Next
    
    Call Grid_Refresh_Checkbox(objGridColunas)

End Sub

Private Function GerarRotinas(ByVal colColunasTabelas As Collection) As Long

Dim lErro As Long
Dim sNL As String
Dim sRotinaLe As String
Dim sRotinaGrava As String
Dim sRotinaExclui As String
Dim sDicErros As String
Dim sDicRotinas As String
Dim sErroGerUlt As String
Dim sScriptRotina As String

On Error GoTo Erro_GerarRotinas

    sNL = Chr(10)
    
    sScriptRotina = "'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA"
    sDicErros = "--SCRIPT CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA"
    
    If optLeitura.Value = vbChecked Then
        lErro = RotinaLe_Cria(colColunasTabelas, sRotinaLe, sDicErros, sDicRotinas)
        If lErro <> SUCESSO Then gError 131717
    
        ScriptLeitura.Text = sScriptRotina & sRotinaLe
    End If

    'Se é uma tabela
    If gsTipoArquivo = ARQUIVO_TABELA Then
    
        If optExclusao.Value = vbChecked Then
            lErro = RotinaExclui_Cria(colColunasTabelas, sRotinaExclui, sDicErros, sDicRotinas)
            If lErro <> SUCESSO Then gError 131718
        
            ScriptExclusao.Text = sScriptRotina & sRotinaExclui
        End If
    
        If optGravacao.Value = vbChecked Then
            lErro = RotinaGrava_Cria(colColunasTabelas, sRotinaGrava, sDicErros, sDicRotinas)
            If lErro <> SUCESSO Then gError 131719
        
            ScriptGravacao.Text = sScriptRotina & sRotinaGrava
        End If
    
    Else
    
        If optExclusao.Value = vbChecked Or optGravacao.Value = vbChecked Then gError 131915
        
    End If
    
    If optDic.Value = vbChecked Then
        lErro = Telas_Cria(sDicRotinas, colColunasTabelas)
        If lErro <> SUCESSO Then gError 131720
        
        ScriptDic.Text = sDicErros & sNL & sNL & sDicRotinas
    End If
    
    GerarRotinas = SUCESSO
    
    Exit Function

Erro_GerarRotinas:

    GerarRotinas = gErr

    Select Case gErr
    
        Case 131715 To 131720
        
        Case 131915
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_VIEW", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143902)

    End Select
    
    Exit Function
    
End Function

Private Sub BotaoExpBrowse_Click()

On Error GoTo Erro_BotaoExpBrowse

    Open CurDir & "\" & NomeTela.Text & "Browse.txt" For Output As #2
    
    Print #2, ScriptBrowse.Text
    
    Close #2
    
    MsgBox "O Arquivo foi exportado para " & CurDir & " com o nome de " & NomeTela.Text & "Browse.txt", vbOKOnly, "SGE"
    
    Exit Sub
    
Erro_BotaoExpBrowse:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143903)

    End Select
    
    Close #2
    
    Exit Sub
    
End Sub

Private Sub BotaoExpType_Click()

On Error GoTo Erro_BotaoExpType

    Open CurDir & "\" & NomeTela.Text & "Type.txt" For Output As #2
    
    Print #2, ScriptType.Text
    
    Close #2
    
    MsgBox "O Arquivo foi exportado para " & CurDir & " com o nome de " & NomeTela.Text & "Type.txt", vbOKOnly, "SGE"
    
    Exit Sub
    
Erro_BotaoExpType:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143904)

    End Select
    
    Close #2
    
    Exit Sub
    
End Sub

Private Sub BotaoExpLeitura_Click()

On Error GoTo Erro_BotaoExpLeitura

    Open CurDir & "\" & NomeTela.Text & "Leitura.txt" For Output As #2
    
    Print #2, ScriptLeitura.Text
    
    Close #2
    
    MsgBox "O Arquivo foi exportado para " & CurDir & " com o nome de " & NomeTela.Text & "Leitura.txt", vbOKOnly, "SGE"
    
    Exit Sub
    
Erro_BotaoExpLeitura:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143905)

    End Select
    
    Close #2
    
    Exit Sub
    
End Sub

Private Sub BotaoExpGravacao_Click()

On Error GoTo Erro_BotaoExpGravacao

    Open CurDir & "\" & NomeTela.Text & "Gravacao.txt" For Output As #2
    
    Print #2, ScriptGravacao.Text
    
    Close #2
    
    MsgBox "O Arquivo foi exportado para " & CurDir & " com o nome de " & NomeTela.Text & "Gravacao.txt", vbOKOnly, "SGE"
    
    Exit Sub
    
Erro_BotaoExpGravacao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143906)

    End Select
    
    Close #2
    
    Exit Sub
    
End Sub

Private Sub BotaoExpExclusao_Click()

On Error GoTo Erro_BotaoExpExclusao

    Open CurDir & "\" & NomeTela.Text & "Exclusao.txt" For Output As #2
    
    Print #2, ScriptExclusao.Text
    
    Close #2
    
    MsgBox "O Arquivo foi exportado para " & CurDir & " com o nome de " & NomeTela.Text & "Exclusao.txt", vbOKOnly, "SGE"
    
    Exit Sub
    
Erro_BotaoExpExclusao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143907)

    End Select
    
    Close #2
    
    Exit Sub
    
End Sub

Private Sub BotaoExpDic_Click()

On Error GoTo Erro_BotaoExpDic

    Open CurDir & "\" & NomeTela.Text & "Dic.txt" For Output As #2
    
    Print #2, ScriptDic.Text
    
    Close #2
    
    MsgBox "O Arquivo foi exportado para " & CurDir & " com o nome de " & NomeTela.Text & "Dic.txt", vbOKOnly, "SGE"
   
    Exit Sub
    
Erro_BotaoExpDic:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143908)

    End Select
    
    Close #2
    
    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEvento = New AdmEvento
    Set gobjTela = New ClassCriaTela
    
    gsErroLeitura = ""
    gbCriarScriptDelete = False
    iCont = 0
    iContAdicao = 0

    Set objGridColunas = New AdmGrid

    Call Carrega_ComboArquivo(NomeArq)
    
    lErro = Inicializa_Grid_Colunas(objGridColunas)
    If lErro <> SUCESSO Then gError 131720
    
    iAlterado = 0
    iAlteradoNomeArq = 0
   
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 131720

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143909)

    End Select

    iAlterado = 0
    iAlteradoNomeArq = 0
    
    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEvento = Nothing
    Set gobjTela = Nothing

    'Liberar as variaveis globais
    Set objGridColunas = Nothing

End Sub

Private Function Preenche_GridColunas(ByVal colColunasTabelas As Collection) As Long

Dim lErro As Long
Dim objColunasTabelas As ClassColunasTabelas
Dim iIndice As Integer

On Error GoTo Erro_Preenche_GridColunas
    
    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridColunas)
    
    iIndice = 0
    
    For Each objColunasTabelas In colColunasTabelas
    
        gsTipoArquivo = Trim(objColunasTabelas.sArquivoTipo)
    
        iIndice = iIndice + 1
    
        GridColunas.TextMatrix(iIndice, iGrid_Browse_Col) = CStr(objColunasTabelas.iBrowse)
        GridColunas.TextMatrix(iIndice, iGrid_Coluna_Col) = objColunasTabelas.sColuna
        GridColunas.TextMatrix(iIndice, iGrid_Descricao_Col) = objColunasTabelas.sColuna
        GridColunas.TextMatrix(iIndice, iGrid_Ordem_Col) = CStr(iIndice)
        GridColunas.TextMatrix(iIndice, iGrid_Tamanho_Col) = CStr(objColunasTabelas.lColunaTamanho)
        GridColunas.TextMatrix(iIndice, iGrid_Precisao_Col) = CStr(objColunasTabelas.lColunaPrecisao)
        GridColunas.TextMatrix(iIndice, iGrid_Tipo_Col) = objColunasTabelas.sColunaTipo
        GridColunas.TextMatrix(iIndice, iGrid_AtribClasse_Col) = objColunasTabelas.sAtributoClasse
        GridColunas.TextMatrix(iIndice, iGrid_TamanhoTela_Col) = CStr(objColunasTabelas.lTamanhoTela)
    
        'Tenta encontrar o subtipo
        Select Case UCase(objColunasTabelas.sColuna)
        
            Case "PRODUTO"
                GridColunas.TextMatrix(iIndice, iGrid_SubTipo_Col) = "6"
            
            Case "CCL"
                GridColunas.TextMatrix(iIndice, iGrid_SubTipo_Col) = "2"
            
            Case "CONTA", "CONTACONTABIL"
                GridColunas.TextMatrix(iIndice, iGrid_SubTipo_Col) = "1"
            
            Case "EXERCICIO"
                GridColunas.TextMatrix(iIndice, iGrid_SubTipo_Col) = "3"
            
            Case "PERIODO"
                GridColunas.TextMatrix(iIndice, iGrid_SubTipo_Col) = "4"
            
            Case "PERCENTUAL"
                GridColunas.TextMatrix(iIndice, iGrid_SubTipo_Col) = "5"
            
            Case "NATUREZA"
                GridColunas.TextMatrix(iIndice, iGrid_SubTipo_Col) = "9"
            
            Case "APROPRIACAO"
                GridColunas.TextMatrix(iIndice, iGrid_SubTipo_Col) = "16"
            
            Case "HORA"
                GridColunas.TextMatrix(iIndice, iGrid_SubTipo_Col) = "26"
            
            Case "TIPOFRETE"
                GridColunas.TextMatrix(iIndice, iGrid_SubTipo_Col) = "31"
            
            Case Else
                GridColunas.TextMatrix(iIndice, iGrid_SubTipo_Col) = "0"
        
        End Select
    
    Next
    
    Call Grid_Refresh_Checkbox(objGridColunas)
            
    'Atualiza o número de linhas existentes
    objGridColunas.iLinhasExistentes = iIndice
            
    Preenche_GridColunas = SUCESSO
    
    Exit Function
    
Erro_Preenche_GridColunas:

    Preenche_GridColunas = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143910)
    
    End Select
    
    Exit Function
    
End Function

Private Sub BotaoGerarCodigo_Click()

Dim lErro As Long
Dim colColunasTabelas As New Collection
Dim sMsg As String
Dim sNL As String

On Error GoTo Erro_BotaoGerarCodigo_Click

    GL_objMDIForm.MousePointer = vbHourglass
   
    sNL = Chr(10)

    lErro = Critica_Tela()
    If lErro <> SUCESSO Then gError 131712

    lErro = Move_Tela_Memoria(colColunasTabelas)
    If lErro <> SUCESSO Then gError 131713

    If optBrowse.Value = vbChecked Then Call GerarBrowse(colColunasTabelas)
    If optType.Value = vbChecked Then Call GerarType(colColunasTabelas)
    If optClasse.Value = vbChecked Then Call GerarClasse(colColunasTabelas)
    Call GerarRotinas(colColunasTabelas)
    If optTela.Value = vbChecked Then Call GerarTela(colColunasTabelas)
    
    If optBrowse.Value = vbChecked Then sMsg = sMsg & "O Script do Browse deve ser rodado no BD Dic." & sNL
    If optClasse.Value = vbChecked Then sMsg = sMsg & "A Classe " & Classe.Text & ".cls foi criada no diretório " & CurDir & sNL
    If optDic.Value = vbChecked Then sMsg = sMsg & "Deve ser rodado o Script do DIC no BD Dic." & sNL
    If optExclusao.Value = vbChecked Then sMsg = sMsg & "A Rotina de Exclusao deve ser colocada na Classe Class" & ModuloClasse.Text & "Grava em Rotinas" & ModuloClasse & sNL
    If optGravacao.Value = vbChecked Then sMsg = sMsg & "A Rotina de Gravação deve ser colocada na Classe Class" & ModuloClasse.Text & "Grava em Rotinas" & ModuloClasse & sNL
    If optLeitura.Value = vbChecked Then sMsg = sMsg & "A Rotina de Leitura deve ser colocada na Classe Class" & ModuloClasse.Text & "Select em Rotinas" & ModuloClasse & sNL
    If optTela.Value = vbChecked Then sMsg = sMsg & "A Tela " & NomeTela.Text & ".ctl foi criada no diretório " & CurDir & sNL
    If optType.Value = vbChecked Then sMsg = sMsg & "O Type deve ser colocado no arquivo .bas." & sNL
    
    If Len(Trim(sMsg)) > 0 Then MsgBox sMsg, vbOKOnly, "SGE"

    ErroGerado.Caption = UltimoErro
    
    iCont = 0
    iContAdicao = 0
    gsErroLeitura = ""

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoGerarCodigo_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 131712 To 131713

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143911)

    End Select
    
    Exit Sub
        
End Sub

Private Sub BotaoTeste_Click()

Dim lErro As Long
Dim obj1 As Object
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoTeste_Click

    If Len(Trim(NomeBrowse.Text)) = 0 Then Exit Sub
    
    Call Chama_Tela(NomeBrowse.Text, colSelecao, obj1, objEvento)

    Exit Sub

Erro_BotaoTeste_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143912)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoVerificar_Click()

Dim lErro As Long
Dim colBrowseArquivo As New Collection
Dim objBrowseArquivo As AdmBrowseArquivo
Dim colColunasTabelas As Collection
Dim colCampo As Collection
Dim sScriptCorrecao As String
Dim sScript As String
Dim sNL As String
Dim sMsg As String
Dim i As Integer
Dim colTabelas As New Collection
Dim sData As String
Dim sNomeCorrecao As String
Dim sNomeErros As String

On Error GoTo Erro_BotaoVerificar_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    sNL = Chr(10)
    
    sData = FormataCpoNum(Year(Now), 4) & FormataCpoNum(Month(Now), 2) & FormataCpoNum(Day(Now), 2) & "_" & FormataCpoNum(Hour(Now), 2) & FormataCpoNum(Minute(Now), 2)

    sNomeCorrecao = CurDir & "\" & "BROWSE_CORRECAO_" & sData & ".sql"
    sNomeErros = CurDir & "\" & "BROWSE_ERROS_" & sData & ".txt"

    Open sNomeErros For Output As #4
    Open sNomeCorrecao For Output As #5

    If optApenasBrowse.Value = vbChecked Then
    
        sScriptCorrecao = ""
    
        If Len(Trim(NomeBrowse.Text)) = 0 Then gError 131910
        
        Set objBrowseArquivo = New AdmBrowseArquivo
    
        lErro = CF("BrowseArquivo_Le", NomeBrowse.Text, objBrowseArquivo)
        If lErro <> SUCESSO Then gError 131911
        
        If InStr(1, objBrowseArquivo.sNomeArq, Chr(0)) <> 0 Then gError 131912
    
        Set colColunasTabelas = New Collection
        Set colCampo = New Collection
    
        lErro = ColunasTabelas_Le(objBrowseArquivo.sNomeArq, colColunasTabelas)
        If lErro <> SUCESSO Then gError 131901
        
        lErro = Campos_Le_Todos2(objBrowseArquivo.sNomeArq, colCampo)
        If lErro <> SUCESSO Then gError 131902
        
        lErro = Valida_Browse(objBrowseArquivo, colColunasTabelas, colCampo, sScript, colTabelas)
        If lErro <> SUCESSO Then gError 131903
        
        If Len(Trim(sScript)) > 0 Then
            sScriptCorrecao = "-- BROWSE " & objBrowseArquivo.sNomeTela
            sScriptCorrecao = sScriptCorrecao & sNL & sScript
        End If
    
        If Len(Trim(sScriptCorrecao)) > 0 Then
            Print #5, sScriptCorrecao
        End If
    
    Else

        lErro = BrowseArquivo_Le_Todos(colBrowseArquivo)
        If lErro <> SUCESSO Then gError 131900
        
        PB.Max = colBrowseArquivo.Count
        
        i = 0
        
        For Each objBrowseArquivo In colBrowseArquivo
        
            i = i + 1
        
            sScriptCorrecao = ""
       
            Set colColunasTabelas = New Collection
            Set colCampo = New Collection
        
            lErro = ColunasTabelas_Le(objBrowseArquivo.sNomeArq, colColunasTabelas)
            If lErro <> SUCESSO Then gError 131901
            
            lErro = Campos_Le_Todos2(objBrowseArquivo.sNomeArq, colCampo)
            If lErro <> SUCESSO Then gError 131902
            
            sScript = ""
            
            lErro = Valida_Browse(objBrowseArquivo, colColunasTabelas, colCampo, sScript, colTabelas)
            If lErro <> SUCESSO Then gError 131903
        
            If Len(Trim(sScript)) > 0 Then
                sScriptCorrecao = "-- BROWSE " & objBrowseArquivo.sNomeTela
                sScriptCorrecao = sScriptCorrecao & sNL & sScript
            End If
        
            If Len(Trim(sScriptCorrecao)) > 0 Then
                Print #5, sScriptCorrecao
            End If
        
            PB.Value = i
        
        Next
                
    End If
    
    Close #5
    Close #4

    GL_objMDIForm.MousePointer = vbDefault
    
    sMsg = "A Validação foi feita com sucesso." & sNL & "O relatório de erros está em " & sNomeErros
    
    If optCorrecao.Value = vbChecked Then
        sMsg = sMsg & sNL & "O arquivo com as correções está em " & sNomeCorrecao
    End If
    
    MsgBox sMsg, vbOKOnly, "SGE"

    PB.Value = 0

    Exit Sub

Erro_BotaoVerificar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 131900 To 131903
        
        Case 131910
            Call Rotina_Erro(vbOKOnly, "ERRO_BROWSE_NAO_PREENCHIDO", gErr)
            NomeBrowse.SetFocus
            
        Case 131911
        
        Case 131912
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ARQUIVO_BROWSEARQUIVO", gErr, NomeBrowse.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143913)

    End Select
    
    Close #5
    Close #4
    
    Exit Sub
    
End Sub

Private Function Valida_Browse(ByVal objBrowseArquivo As AdmBrowseArquivo, ByVal colColunasTabelas As Collection, ByVal colCampo As Collection, sScriptCorrecao As String, ByVal colTabelas As Collection)

Dim lErro As Long
Dim objColunasTabelas As ClassColunasTabelas
Dim objCampos As AdmCampos
Dim objCamposAux As AdmCampos
Dim bAchou As Boolean
Dim bErro As Boolean
Dim iSeqErro As Integer
Dim iSeq As Integer
Dim sNL As String
Dim sErro As String
Dim bExisteTabela As Boolean
Dim bTabelaJaVerificada As Boolean
Dim iPos As Integer
Dim iPosAux As Integer
Dim objBrowseParamSelecao As AdmBrowseParamSelecao
Dim colParamSelecao As New Collection
Dim bAchouFilialEmpresa As Boolean
Dim iCount As Integer

On Error GoTo Erro_Valida_Browse

    iContAdicao = 0

    sNL = Chr(10)
    bErro = False
    bExisteTabela = True
     
    Print #4, ""
    Print #4, "VERIFICAÇÃO DO BROWSE " & objBrowseArquivo.sNomeTela
    Print #4, ""
    
    iSeqErro = 0
    
    'Verifica se a tabela já foi verificada anteriormente
    bTabelaJaVerificada = False
    If colCampo.Count > 0 Then
    
        For Each objCamposAux In colTabelas
        
            If objCamposAux.sNomeArq = colCampo.Item(1).sNomeArq Then
                bTabelaJaVerificada = True
                Exit For
            End If
        
        Next
    
    End If
     
    If colColunasTabelas.Count = 0 Then
     
        bErro = True
        iSeqErro = iSeqErro + 1
        If iSeqErro = 1 Then
            iSeq = iSeq + 1
            Print #4, TECLA_TAB & CStr(iSeq) & SEPARADOR & "ERROS REFERENTES A TABELA."
        End If
        sErro = CStr(iSeqErro) & SEPARADOR & "A Tabela/View " & objBrowseArquivo.sNomeArq & " não existe no SGEDados."
        Print #4, TECLA_TAB & TECLA_TAB & sErro
        
        bExisteTabela = False
        
        If optCorrecao.Value = vbChecked And Not bTabelaJaVerificada Then
            sScriptCorrecao = sScriptCorrecao & sNL & "-- O Erro '" & sErro & "' não pode ser corrigido." & sNL
        End If
        
    End If
     
    iSeqErro = 0

    'Classe do objeto
    lErro = Critica_Objeto(objBrowseArquivo.sProjetoObjeto & "." & objBrowseArquivo.sClasseObjeto)
    If lErro <> SUCESSO Then
       
        bErro = True
        iSeqErro = iSeqErro + 1
        If iSeqErro = 1 Then
            iSeq = iSeq + 1
            Print #4, TECLA_TAB & CStr(iSeq) & SEPARADOR & "ERROS REFERENTES A REGISTROS NA TABELA BROWSEARQUIVO."
        End If
        
        sErro = CStr(iSeqErro) & SEPARADOR & "ERRO: A Classe " & objBrowseArquivo.sClasseObjeto & " não pertence ao projeto " & objBrowseArquivo.sProjetoObjeto
        Print #4, TECLA_TAB & TECLA_TAB & sErro
        
        If optCorrecao.Value = vbChecked Then
            sScriptCorrecao = sScriptCorrecao & sNL & "--O Erro '" & sErro & "' não pode ser corrigido." & sNL
        End If
        
    End If
    
    'Classe formata
    lErro = Critica_Objeto(objBrowseArquivo.sProjeto & "." & objBrowseArquivo.sClasse)
    If lErro <> SUCESSO Then
       
        bErro = True
        iSeqErro = iSeqErro + 1
        If iSeqErro = 1 Then
            iSeq = iSeq + 1
            Print #4, TECLA_TAB & CStr(iSeq) & SEPARADOR & "ERROS REFERENTES A REGISTROS NA TABELA BROWSEARQUIVO."
        End If
        
        sErro = CStr(iSeqErro) & SEPARADOR & "ERRO: A Classe " & objBrowseArquivo.sClasse & " não pertence ao projeto " & objBrowseArquivo.sProjeto
        Print #4, TECLA_TAB & TECLA_TAB & sErro
        
        If optCorrecao.Value = vbChecked Then
            sScriptCorrecao = sScriptCorrecao & sNL & "--O Erro '" & sErro & "' não pode ser corrigido." & sNL
        End If
        
    End If
    
    'Se não existe a tabela => Termina a validação
    If Not bExisteTabela Then
        Valida_Browse = SUCESSO
        Exit Function
    End If
    
    'Verifica se o campo SQLSelecao Está Correto
    If Len(Trim(objBrowseArquivo.sSelecaoSQL)) > 0 Then
    
        lErro = Tabela_Le_Generico(objBrowseArquivo.sNomeArq, objBrowseArquivo.sSelecaoSQL)
        If lErro <> SUCESSO Then
        
            bErro = True
            iSeqErro = iSeqErro + 1
            If iSeqErro = 1 Then
                iSeq = iSeq + 1
                Print #4, TECLA_TAB & CStr(iSeq) & SEPARADOR & "ERROS REFERENTES A REGISTROS NA TABELA BROWSEARQUIVO."
            End If
            
            sErro = CStr(iSeqErro) & SEPARADOR & "ERRO: O Filtro '" & objBrowseArquivo.sSelecaoSQL & "' em SQLSelecao não está correto."
            Print #4, TECLA_TAB & TECLA_TAB & sErro
            
            If optCorrecao.Value = vbChecked Then
                sScriptCorrecao = sScriptCorrecao & sNL & "--O Erro '" & sErro & "' não pode ser corrigido." & sNL
            End If
        
        End If
        
    End If
    
    iSeqErro = 0
    
'    'Se tem a FilialEmpresa no Filtro ela deve estar no ParamSelecao
'    iPos = InStr(1, UCase(objBrowseArquivo.sSelecaoSQL), "FILIALEMPRESA")
'    If iPos <> 0 Then
'
'        lErro = BrowseParamSelecao_Le(objBrowseArquivo.sNomeTela, colParamSelecao)
'        If lErro <> SUCESSO Then gError 131905
'
'        bAchouFilialEmpresa = False
'
'        For Each objBrowseParamSelecao In colParamSelecao
'
'            If UCase(objBrowseParamSelecao.sProperty) = "GIFILIALEMPRESA" Then
'                bAchouFilialEmpresa = True
'                Exit For
'            End If
'
'        Next
'
'        If Not bAchouFilialEmpresa Then
'
'            iCount = 0
'
'            iPosAux = InStr(1, Left(UCase(objBrowseArquivo.sSelecaoSQL), iPos), "?")
'
'            Do While iPosAux <> 0
'                iCount = iCount + 1
'
'                iPosAux = InStr(iPosAux + 1, Left(UCase(objBrowseArquivo.sSelecaoSQL), iPos), "?")
'            Loop
'
'            bErro = True
'            iSeqErro = iSeqErro + 1
'            If iSeqErro = 1 Then
'                iSeq = iSeq + 1
'                Print #4, TECLA_TAB & CStr(iSeq) & SEPARADOR & "ERROS REFERENTES A REGISTROS NA TABELA BROWSEPARAMSELECAO."
'            End If
'
'            sErro = CStr(iSeqErro) & SEPARADOR & "ERRO: Não existe parâmetro passado para o campo FilialEmpresa contido em SQLSelecao em BrowseArquivo."
'            Print #4, TECLA_TAB & TECLA_TAB & sErro
'
'            If optCorrecao.Value = vbChecked Then
'                sScriptCorrecao = sScriptCorrecao & sNL & "INSERT INTO BrowseParamSelecao (NomeTela, Ordem, Projeto, Classe, Property)" & sNL & "VALUES('" & objBrowseArquivo.sNomeTela & "'," & CStr(iCount) & ",'admlib','Adm','giFilialEmpresa')" & sNL & "GO" & sNL
'            End If
'
'        End If
'
'    End If
'
'    iSeqErro = 0

    'Verifica existência dos campos na tabela
    For Each objCampos In colCampo
    
         bAchou = False
        
        For Each objColunasTabelas In colColunasTabelas
        
            If UCase(objCampos.sNome) = UCase(objColunasTabelas.sColuna) Then
                bAchou = True
                Exit For
            End If
        
        Next
        
        If Not bAchou Then
        
            bErro = True
            iSeqErro = iSeqErro + 1
            If iSeqErro = 1 Then
                iSeq = iSeq + 1
                Print #4, TECLA_TAB & CStr(iSeq) & SEPARADOR & "ERROS REFERENTES A REGISTROS NA TABELA CAMPOS."
            End If
            Print #4, TECLA_TAB & TECLA_TAB & CStr(iSeqErro) & SEPARADOR & "ERRO: O Campo " & objCampos.sNome & " não pertence a tabela " & objBrowseArquivo.sNomeArq
        
            If optCorrecao.Value = vbChecked And Not bTabelaJaVerificada Then
                sScriptCorrecao = sScriptCorrecao & sNL & "DELETE Campos WHERE NomeArq = '" & objBrowseArquivo.sNomeArq & "' AND Nome = '" & objCampos.sNome & "'"
                sScriptCorrecao = sScriptCorrecao & sNL & "GO" & sNL
            End If
        
        End If
    
    Next
    
    If optAviso.Value = vbChecked Then
        
        'Verifica não existência dos campos em campos
        For Each objColunasTabelas In colColunasTabelas
        
             bAchou = False
            
            For Each objCampos In colCampo
            
                If UCase(objCampos.sNome) = UCase(objColunasTabelas.sColuna) Then
                    bAchou = True
                    Exit For
                End If
            
            Next
            
            If Not bAchou Then
            
                bErro = True
                iSeqErro = iSeqErro + 1
                If iSeqErro = 1 Then
                    iSeq = iSeq + 1
                    Print #4, TECLA_TAB & CStr(iSeq) & SEPARADOR & "ERROS REFERENTES A REGISTROS NA TABELA CAMPOS."
                End If
                Print #4, TECLA_TAB & TECLA_TAB & CStr(iSeqErro) & SEPARADOR & "AVISO: O Campo " & objColunasTabelas.sColuna & " não está cadastrado na tabela CAMPOS"
            
                If optCorrecao.Value = vbChecked And Not bTabelaJaVerificada Then
                    sScriptCorrecao = sScriptCorrecao & sNL & Incluir_Coluna_Em_Campos(objColunasTabelas, colCampo)
                    sScriptCorrecao = sScriptCorrecao & sNL & "GO" & sNL
                End If
            
            End If
        
        Next
        
    End If
    
    lErro = Valida_BrowseCampo(objBrowseArquivo, colColunasTabelas, colCampo, iSeq, bErro, sScriptCorrecao)
    If lErro <> SUCESSO Then gError 131905
  
    lErro = Valida_BrowseIndice(objBrowseArquivo, colColunasTabelas, colCampo, iSeq, bErro, sScriptCorrecao)
    If lErro <> SUCESSO Then gError 131905
        
    If Not bErro Then Print #4, TECLA_TAB & TECLA_TAB & "BROWSE SEM ERROS"
    
    If colCampo.Count > 0 Then colTabelas.Add colCampo.Item(1)
    
    Valida_Browse = SUCESSO
    
    Exit Function
    
Erro_Valida_Browse:

    Valida_Browse = gErr
    
    Select Case gErr
    
        Case 131905
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143914)
            
    End Select
        
    Exit Function
    
End Function

Private Function Valida_BrowseCampo(ByVal objBrowseArquivo As AdmBrowseArquivo, ByVal colColunasTabelas As Collection, ByVal colCampo As Collection, iSeq As Integer, bErro As Boolean, sScriptCorrecao As String)

Dim lErro As Long
Dim objColunasTabelas As ClassColunasTabelas
Dim objCampos As AdmCampos
Dim objBrowseCampo As AdmBrowseCampo
Dim colBrowseCampo As New Collection
Dim bAchou As Boolean
Dim iSeqErro As Integer
Dim sErro As String
Dim sNL As String

On Error GoTo Erro_Valida_BrowseCampo

    sNL = Chr(10)

    lErro = CF("BrowseCampo_Le", objBrowseArquivo.sNomeTela, colBrowseCampo)
    If lErro <> SUCESSO Then gError 131906

    'Verifica não existência dos campos em campos
    For Each objBrowseCampo In colBrowseCampo
    
         bAchou = False
        
        For Each objCampos In colCampo
        
            If UCase(objCampos.sNome) = UCase(objBrowseCampo.sNomeCampo) Then
                bAchou = True
                Exit For
            End If
        
        Next
        
        If Not bAchou Then
        
            bErro = True
            iSeqErro = iSeqErro + 1
            If iSeqErro = 1 Then
                iSeq = iSeq + 1
                Print #4, TECLA_TAB & CStr(iSeq) & SEPARADOR & "ERROS REFERENTES A REGISTROS NA TABELA BROWSECAMPO."
            End If
            Print #4, TECLA_TAB & TECLA_TAB & CStr(iSeqErro) & SEPARADOR & "ERRO: O Campo " & objBrowseCampo.sNomeCampo & " não está cadastrado na tabela CAMPOS"
        
            If optCorrecao.Value = vbChecked Then
                sScriptCorrecao = sScriptCorrecao & Verificar_Existencia_Campo_BrowseCampo(objBrowseArquivo, colColunasTabelas, objBrowseCampo.sNomeCampo)
            End If
        
        End If
        
        lErro = Critica_Objeto(objBrowseArquivo.sProjetoObjeto & "." & objBrowseArquivo.sClasseObjeto)
        If lErro = SUCESSO Then
            lErro = Critica_ObjetoAtributo(objBrowseArquivo.sProjetoObjeto & "." & objBrowseArquivo.sClasseObjeto, objBrowseCampo.sNome)
            If lErro <> SUCESSO Then
            
                bErro = True
                iSeqErro = iSeqErro + 1
                If iSeqErro = 1 Then
                    iSeq = iSeq + 1
                    Print #4, TECLA_TAB & CStr(iSeq) & SEPARADOR & "ERROS REFERENTES A REGISTROS NA TABELA BROWSECAMPO."
                End If
                
                sErro = CStr(iSeqErro) & SEPARADOR & "ERRO: O Atributo " & objBrowseCampo.sNome & " não pertence a classe " & objBrowseArquivo.sClasseObjeto & " em " & objBrowseArquivo.sProjetoObjeto
                Print #4, TECLA_TAB & TECLA_TAB & sErro
            
                If optCorrecao.Value = vbChecked Then
                    sScriptCorrecao = sScriptCorrecao & sNL & "--O Erro '" & sErro & "' não pode ser resolvido." & sNL
                End If
            
            End If
        End If
    
    Next
    
    lErro = Valida_GrupoBrowseCampo(objBrowseArquivo, colColunasTabelas, colBrowseCampo, colCampo, iSeq, bErro, sScriptCorrecao)
    If lErro <> SUCESSO Then gError 131907
        
    Valida_BrowseCampo = SUCESSO
    
    Exit Function
    
Erro_Valida_BrowseCampo:

    Valida_BrowseCampo = gErr
    
    Select Case gErr
    
        Case 131906 To 131907
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143915)
            
    End Select
        
    Exit Function
    
End Function

Private Function Valida_BrowseIndice(ByVal objBrowseArquivo As AdmBrowseArquivo, ByVal colColunasTabelas As Collection, ByVal colCampo As Collection, iSeq As Integer, bErro As Boolean, sScriptCorrecao As String)

Dim lErro As Long
Dim objColunasTabelas As ClassColunasTabelas
Dim objBrowseIndice As AdmBrowseIndice
Dim colBrowseIndice As New Collection
Dim bAchou As Boolean
Dim iSeqErro As Integer
Dim sErro As String
Dim sNL As String
Dim vCampoIndice As Variant
Dim vCampoIndiceAux As Variant
Dim colCampoIndice As New Collection
Dim sAux As String
Dim iPos As Integer
Dim colUsuarios As New Collection
Dim objUsuario As ClassUsuarios

On Error GoTo Erro_Valida_BrowseIndice

    sNL = Chr(10)

    lErro = CF("BrowseIndice_Le", objBrowseArquivo.sNomeTela, colBrowseIndice)
    If lErro <> SUCESSO Then gError 131906

    For Each objBrowseIndice In colBrowseIndice
    
        iPos = InStr(1, objBrowseIndice.sOrdenacaoSQL, ",")
        sAux = objBrowseIndice.sOrdenacaoSQL
        
        Do While iPos <> 0
        
            vCampoIndice = Trim(left(sAux, iPos - 1))
            sAux = right(sAux, Len(sAux) - iPos)
            iPos = InStr(1, sAux, ",")
            
            bAchou = False
            For Each vCampoIndiceAux In colCampoIndice
            
                If UCase(vCampoIndice) = UCase(vCampoIndiceAux) Then
                    bAchou = True
                    Exit For
                End If
            
            Next
            
            If Not bAchou Then
                colCampoIndice.Add vCampoIndice
            End If
        
        Loop
     
    Next

    For Each vCampoIndice In colCampoIndice

        bAchou = False
        For Each objColunasTabelas In colColunasTabelas
        
            If UCase(objColunasTabelas.sColuna) = UCase(vCampoIndice) Then
                bAchou = True
                Exit For
            End If
        
        Next
        
        If Not bAchou Then
            
            bErro = True
            iSeqErro = iSeqErro + 1
            If iSeqErro = 1 Then
                iSeq = iSeq + 1
                Print #4, TECLA_TAB & CStr(iSeq) & SEPARADOR & "ERROS REFERENTES A REGISTROS NA TABELA BROWSEINDICE."
            End If
            
            sErro = CStr(iSeqErro) & SEPARADOR & "ERRO: O Campo '" & vCampoIndice & "' não pertence a tabela '" & objBrowseArquivo.sNomeArq & "'."
            Print #4, TECLA_TAB & TECLA_TAB & sErro
        
            If optCorrecao.Value = vbChecked Then
                sScriptCorrecao = sScriptCorrecao & sNL & "--O Erro '" & sErro & "' não pode ser resolvido." & sNL
            End If
                
        End If
                
    Next
    
    'BROWSEINDICEUSUARIO
    iSeqErro = 0
    
    lErro = CF("Usuarios_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then gError 131906
    
    For Each objUsuario In colUsuarios
    
        Set colCampoIndice = New Collection
    
        lErro = CF("BrowseIndiceUsuario_Le", objBrowseArquivo.sNomeTela, objUsuario.sCodUsuario, colBrowseIndice)
        If lErro <> SUCESSO Then gError 131906
    
        For Each objBrowseIndice In colBrowseIndice
        
            iPos = InStr(1, objBrowseIndice.sOrdenacaoSQL, ",")
            sAux = objBrowseIndice.sOrdenacaoSQL
            
            Do While iPos <> 0
            
                vCampoIndice = Trim(left(sAux, iPos - 1))
                sAux = right(sAux, Len(sAux) - iPos)
                iPos = InStr(1, sAux, ",")
                
                bAchou = False
                For Each vCampoIndiceAux In colCampoIndice
                
                    If UCase(vCampoIndice) = UCase(vCampoIndiceAux) Then
                        bAchou = True
                        Exit For
                    End If
                
                Next
                
                If Not bAchou Then
                    colCampoIndice.Add vCampoIndice
                End If
            
            Loop
         
        Next
    
        For Each vCampoIndice In colCampoIndice
    
            bAchou = False
            For Each objColunasTabelas In colColunasTabelas
            
                If UCase(objColunasTabelas.sColuna) = UCase(vCampoIndice) Then
                    bAchou = True
                    Exit For
                End If
            
            Next
            
            If Not bAchou Then
                
                bErro = True
                iSeqErro = iSeqErro + 1
                If iSeqErro = 1 Then
                    iSeq = iSeq + 1
                    Print #4, TECLA_TAB & CStr(iSeq) & SEPARADOR & "ERROS REFERENTES A REGISTROS NA TABELA BROWSEINDICEUSUARIO."
                End If
                
                sErro = CStr(iSeqErro) & SEPARADOR & "ERRO: O Campo '" & vCampoIndice & "' no índice do usuário '" & objUsuario.sCodUsuario & "' não pertence a tabela '" & objBrowseArquivo.sNomeArq & "'."
                Print #4, TECLA_TAB & TECLA_TAB & sErro
            
                If optCorrecao.Value = vbChecked Then
                    sScriptCorrecao = sScriptCorrecao & sNL & "--O Erro '" & sErro & "' não pode ser resolvido." & sNL
                End If
                    
            End If
                    
        Next
        
    Next
    
    Valida_BrowseIndice = SUCESSO
    
    Exit Function
    
Erro_Valida_BrowseIndice:

    Valida_BrowseIndice = gErr
    
    Select Case gErr
    
        Case 131906 To 131907
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143916)
            
    End Select
        
    Exit Function
    
End Function

Private Function Valida_BrowseUsuarioCampo(ByVal objBrowseArquivo As AdmBrowseArquivo, ByVal colColunasTabelas As Collection, ByVal colGrupoBrowseCampo As Collection, ByVal colCampo As Collection, iSeq As Integer, bErro As Boolean, sScriptCorrecao As String)

Dim lErro As Long
Dim objGrupoBrowseCampo As AdmGrupoBrowseCampo
Dim colBrowseUsuarioCampo As New Collection
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
Dim iSeqErro As Integer
Dim bAchou As Boolean
Dim sNL As String
Dim colSaida As New Collection
Dim colCampos As New Collection
Dim sUsuarioAnt As String
Dim iPosicaoAnt As Integer
Dim sErro As String

On Error GoTo Erro_Valida_BrowseUsuarioCampo

    sNL = Chr(10)

    lErro = BrowseUsuarioCampo_Le_Todos(objBrowseArquivo.sNomeTela, colBrowseUsuarioCampo)
    If lErro <> SUCESSO Then gError 131909
    
    'Verifica não existência dos campos em campos
    For Each objBrowseUsuarioCampo In colBrowseUsuarioCampo
    
         bAchou = False
        
        For Each objGrupoBrowseCampo In colGrupoBrowseCampo
        
            If UCase(objGrupoBrowseCampo.sNome) = UCase(objBrowseUsuarioCampo.sNome) Then
                bAchou = True
                Exit For
            End If
        
        Next
        
        If Not bAchou Then
        
            bErro = True
            iSeqErro = iSeqErro + 1
            If iSeqErro = 1 Then
                iSeq = iSeq + 1
                Print #4, TECLA_TAB & CStr(iSeq) & SEPARADOR & "ERROS REFERENTES A REGISTROS NA TABELA BROWSEUSUARIOCAMPO."
            End If
            Print #4, TECLA_TAB & TECLA_TAB & CStr(iSeqErro) & SEPARADOR & "ERRO: O Campo " & objBrowseUsuarioCampo.sNome & " não está cadastrado na tabela GRUPOBROWSECAMPO"
        
            If optCorrecao.Value = vbChecked Then
                sScriptCorrecao = sScriptCorrecao & Verificar_Existencia_Campo_BrowseUsuarioCampo(objBrowseArquivo, colColunasTabelas, objBrowseUsuarioCampo.sNome)
            End If
        
        End If
    
    Next
    
    'Verifica Ordenação dos Campos
    colCampos.Add "sCodUsuario"
    colCampos.Add "iPosicaoTela"
    
    Call Ordena_Colecao(colBrowseUsuarioCampo, colSaida, colCampos)
    
    For Each objBrowseUsuarioCampo In colSaida
    
        If sUsuarioAnt <> objBrowseUsuarioCampo.sCodUsuario Then
            sUsuarioAnt = objBrowseUsuarioCampo.sCodUsuario
            iPosicaoAnt = 0
        End If
        
        If iPosicaoAnt + 1 <> objBrowseUsuarioCampo.iPosicaoTela Then
        
            bErro = True
            iSeqErro = iSeqErro + 1
            If iSeqErro = 1 Then
                iSeq = iSeq + 1
                Print #4, TECLA_TAB & CStr(iSeq) & SEPARADOR & "ERROS REFERENTES A REGISTROS NA TABELA BROWSEUSUARIOCAMPO."
            End If
            sErro = CStr(iSeqErro) & SEPARADOR & "ERRO: A Coluna 'PosicaoTela' não está na ordem correta em '" & objBrowseUsuarioCampo.sNome & "' para o usuário '" & objBrowseUsuarioCampo.sCodUsuario & "'."
            Print #4, TECLA_TAB & TECLA_TAB & sErro
        
            If optCorrecao.Value = vbChecked Then
                sScriptCorrecao = sScriptCorrecao & sNL & "-- O Erro '" & sErro & "' não pode ser corrigido." & sNL
            End If
            
        End If
    
        iPosicaoAnt = objBrowseUsuarioCampo.iPosicaoTela
           
    Next
    
    Valida_BrowseUsuarioCampo = SUCESSO
    
    Exit Function
    
Erro_Valida_BrowseUsuarioCampo:

    Valida_BrowseUsuarioCampo = gErr
    
    Select Case gErr
    
        Case 131909
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143917)
            
    End Select
        
    Exit Function
    
End Function

Private Function Valida_GrupoBrowseCampo(ByVal objBrowseArquivo As AdmBrowseArquivo, ByVal colColunasTabelas As Collection, ByVal colBrowseCampo As Collection, ByVal colCampo As Collection, iSeq As Integer, bErro As Boolean, sScriptCorrecao As String)

Dim lErro As Long
Dim objCampo As AdmCampos
Dim colGrupoBrowseCampo As New Collection
Dim objGrupoBrowseCampo As AdmGrupoBrowseCampo
Dim iSeqErro As Integer
Dim bAchou As Boolean
Dim sNL As String

On Error GoTo Erro_Valida_GrupoBrowseCampo

    sNL = Chr(10)

    lErro = GrupoBrowseCampo_Le_Todos(objBrowseArquivo.sNomeTela, colGrupoBrowseCampo)
    If lErro <> SUCESSO Then gError 131909
    
    'Verifica não existência dos campos em campos
    For Each objGrupoBrowseCampo In colGrupoBrowseCampo
    
         bAchou = False
        
        For Each objCampo In colCampo
        
            If UCase(objCampo.sNome) = UCase(objGrupoBrowseCampo.sNome) Then
                bAchou = True
                Exit For
            End If
        
        Next
        
        If Not bAchou Then
        
            bErro = True
            iSeqErro = iSeqErro + 1
            If iSeqErro = 1 Then
                iSeq = iSeq + 1
                Print #4, TECLA_TAB & CStr(iSeq) & SEPARADOR & "ERROS REFERENTES A REGISTROS NA TABELA GRUPOBROWSECAMPO."
            End If
            Print #4, TECLA_TAB & TECLA_TAB & CStr(iSeqErro) & SEPARADOR & "ERRO: O Campo " & objGrupoBrowseCampo.sNome & " não está cadastrado na tabela CAMPOS"
        
            If optCorrecao.Value = vbChecked Then
                sScriptCorrecao = sScriptCorrecao & Verificar_Existencia_Campo_GrupoBrowseCampo(objBrowseArquivo, colColunasTabelas, objGrupoBrowseCampo.sNome)
            End If
        
        End If
    
    Next
    
    lErro = Valida_BrowseUsuarioCampo(objBrowseArquivo, colColunasTabelas, colGrupoBrowseCampo, colCampo, iSeq, bErro, sScriptCorrecao)
    If lErro <> SUCESSO Then gError 131908

    Valida_GrupoBrowseCampo = SUCESSO
    
    Exit Function
    
Erro_Valida_GrupoBrowseCampo:

    Valida_GrupoBrowseCampo = gErr
    
    Select Case gErr
    
        Case 131908
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143918)
            
    End Select
        
    Exit Function
    
End Function

Private Function Incluir_Coluna_Em_Campos(ByVal objColunasTabelas As ClassColunasTabelas, ByVal colCampo As Collection) As String

Dim lErro As Long
Dim sScript As String
Dim objCampo As AdmCampos
Dim iMaxOrdinal As Integer
Dim sSigla As String
Dim iTipo As Integer
Dim sTipoVB As String
Dim sNL As String

On Error GoTo Erro_Incluir_Coluna_Em_Campos

    sNL = Chr(10)

    Call ObtemSiglaTipo(objColunasTabelas.sColunaTipo, sSigla, iTipo, sTipoVB)
    
    iContAdicao = iContAdicao + 1
    iMaxOrdinal = iContAdicao + colCampo.Count
    
    With objColunasTabelas
    
        sScript = "INSERT INTO Campos(NomeArq,Nome,Descricao,Obrigatorio,Imexivel, Ativo, Tipo, Tamanho, Precisao, Decimais, TamExibicao, TituloEntradaDados, TituloGrid, Ordinal, Alinhamento, SubTipo )"
        sScript = sScript & sNL & "VALUES ('" & .sArquivo & "','" & .sColuna & "','" & .sColuna & "',1,0,1," & CStr(iTipo) & "," & CStr(.lColunaTamanho) & "," & CStr(.lColunaPrecisao) & "," & CStr(.lColunaPrecisao) & ",0, '" & .sColuna & "', '" & .sColuna & "', " & CStr(iMaxOrdinal) & ",0,0)" & sNL & "GO" & sNL
   
    End With

    Incluir_Coluna_Em_Campos = sScript
    
    Exit Function
    
Erro_Incluir_Coluna_Em_Campos:

    Incluir_Coluna_Em_Campos = "--NÃO FOI POSSÍVEL CONCLUIR ESSE SCRIPT"
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143919)
            
    End Select
        
    Exit Function
    
End Function

Private Function Verificar_Existencia_Campo_BrowseCampo(ByVal objBrowseArquivo As AdmBrowseArquivo, ByVal colColunasTabelas As Collection, ByVal sNomeCampo As String) As String

Dim lErro As Long
Dim sScript As String
Dim objColunasTabelas As ClassColunasTabelas
Dim sNL As String
Dim bAchou As Boolean

On Error GoTo Erro_Verificar_Existencia_Campo_BrowseCampo

    sNL = Chr(10)

    bAchou = False
    For Each objColunasTabelas In colColunasTabelas
    
        If UCase(objColunasTabelas.sColuna) = UCase(sNomeCampo) Then
            bAchou = True
            Exit For
        End If
    
    Next
    
    If Not bAchou Then
        sScript = sNL & "DELETE BrowseCampo" & sNL & "WHERE NomeTela = '" & objBrowseArquivo.sNomeTela & "' AND NomeCampo = '" & sNomeCampo & "'" & sNL & "GO" & sNL
    End If

    Verificar_Existencia_Campo_BrowseCampo = sScript
    
    Exit Function
    
Erro_Verificar_Existencia_Campo_BrowseCampo:

    Verificar_Existencia_Campo_BrowseCampo = sNL & "--NÃO FOI POSSÍVEL CONCLUIR ESSE SCRIPT"
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143920)
            
    End Select
        
    Exit Function
End Function

Private Function Verificar_Existencia_Campo_GrupoBrowseCampo(ByVal objBrowseArquivo As AdmBrowseArquivo, ByVal colColunasTabelas As Collection, ByVal sNomeCampo As String) As String

Dim lErro As Long
Dim sScript As String
Dim objColunasTabelas As ClassColunasTabelas
Dim sNL As String
Dim bAchou As Boolean

On Error GoTo Erro_Verificar_Existencia_Campo_GrupoBrowseCampo

    sNL = Chr(10)

    bAchou = False
    For Each objColunasTabelas In colColunasTabelas
    
        If UCase(objColunasTabelas.sColuna) = UCase(sNomeCampo) Then
            bAchou = True
            Exit For
        End If
    
    Next
    
    If Not bAchou Then
        sScript = sNL & "DELETE GrupoBrowseCampo" & sNL & "WHERE NomeTela = '" & objBrowseArquivo.sNomeTela & "' AND Nome = '" & sNomeCampo & "'" & sNL & "GO" & sNL
    End If

    Verificar_Existencia_Campo_GrupoBrowseCampo = sScript
    
    Exit Function
    
Erro_Verificar_Existencia_Campo_GrupoBrowseCampo:

    Verificar_Existencia_Campo_GrupoBrowseCampo = sNL & "--NÃO FOI POSSÍVEL CONCLUIR ESSE SCRIPT"
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143921)
            
    End Select
        
    Exit Function
End Function

Private Function Verificar_Existencia_Campo_BrowseUsuarioCampo(ByVal objBrowseArquivo As AdmBrowseArquivo, ByVal colColunasTabelas As Collection, ByVal sNomeCampo As String) As String

Dim lErro As Long
Dim sScript As String
Dim objColunasTabelas As ClassColunasTabelas
Dim sNL As String
Dim bAchou As Boolean

On Error GoTo Erro_Verificar_Existencia_Campo_BrowseUsuarioCampo

    sNL = Chr(10)

    bAchou = False
    For Each objColunasTabelas In colColunasTabelas
    
        If UCase(objColunasTabelas.sColuna) = UCase(sNomeCampo) Then
            bAchou = True
            Exit For
        End If
    
    Next
    
    If Not bAchou Then
        sScript = sNL & "DELETE BrowseUsuarioCampo" & sNL & "WHERE NomeTela = '" & objBrowseArquivo.sNomeTela & "' AND Nome = '" & sNomeCampo & "'" & sNL & "GO" & sNL
    Else
        sScript = sNL & "INSERT INTO GrupoBrowseCampo (CodGrupo, NomeTela, NomeArq, Nome)"
        sScript = sScript & sNL & "VALUES ( 'supervisor', '" & objBrowseArquivo.sNomeTela & "', '" & objBrowseArquivo.sNomeArq & "', '" & sNomeCampo & "')" & sNL & "GO" & sNL
    End If

    Verificar_Existencia_Campo_BrowseUsuarioCampo = sScript
    
    Exit Function
    
Erro_Verificar_Existencia_Campo_BrowseUsuarioCampo:

    Verificar_Existencia_Campo_BrowseUsuarioCampo = sNL & "--NÃO FOI POSSÍVEL CONCLUIR ESSE SCRIPT"
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143922)
            
    End Select
        
    Exit Function
    
End Function

Private Sub ModuloAcesso_Click()

    If Len(Trim(ModuloClasse.Text)) = 0 Then
        ModuloClasse.Text = ModuloAcesso.Text
    End If

    If Len(Trim(ModuloFormata.Text)) = 0 Then
        ModuloFormata.Text = ModuloAcesso.Text
    End If

    If Len(Trim(ModuloTela.Text)) = 0 Then
        ModuloTela.Text = ModuloAcesso.Text
    End If

End Sub

Private Sub ModuloClasse_Click()

    If Len(Trim(ModuloAcesso.Text)) = 0 Then
        ModuloAcesso.Text = ModuloClasse.Text
    End If

    If Len(Trim(ModuloFormata.Text)) = 0 Then
        ModuloFormata.Text = ModuloClasse.Text
    End If

    If Len(Trim(ModuloTela.Text)) = 0 Then
        ModuloTela.Text = ModuloClasse.Text
    End If

End Sub

Private Sub ModuloFormata_Click()

    If Len(Trim(ModuloClasse.Text)) = 0 Then
        ModuloClasse.Text = ModuloFormata.Text
    End If

    If Len(Trim(ModuloAcesso.Text)) = 0 Then
        ModuloAcesso.Text = ModuloFormata.Text
    End If

    If Len(Trim(ModuloTela.Text)) = 0 Then
        ModuloTela.Text = ModuloFormata.Text
    End If

End Sub

Private Sub ModuloTela_Click()

    If Len(Trim(ModuloClasse.Text)) = 0 Then
        ModuloClasse.Text = ModuloTela.Text
    End If

    If Len(Trim(ModuloFormata.Text)) = 0 Then
        ModuloFormata.Text = ModuloTela.Text
    End If

    If Len(Trim(ModuloAcesso.Text)) = 0 Then
        ModuloAcesso.Text = ModuloTela.Text
    End If

End Sub

Private Sub NomeArq_Change()

    iAlteradoNomeArq = REGISTRO_ALTERADO

End Sub

Private Sub NomeArq_Click()
    
    iAlteradoNomeArq = REGISTRO_ALTERADO

End Sub

Private Sub NomeArq_Validate(Cancel As Boolean)

Dim lErro As Long
Dim colColunasTabelas As New Collection

On Error GoTo Erro_NomeArq_Validate

    If Len(Trim(NomeArq.Text)) <> 0 And iAlteradoNomeArq = REGISTRO_ALTERADO And UCase(gsNomeArqAnterior) <> UCase(NomeArq.Text) Then
    
        gsNomeArqAnterior = NomeArq.Text
    
        lErro = ColunasTabelas_Le(NomeArq.Text, colColunasTabelas)
        If lErro <> SUCESSO Then gError 131721
        
        lErro = Preenche_GridColunas(colColunasTabelas)
        If lErro <> SUCESSO Then gError 131722
        
        NomeBrowse.Text = NomeArq.Text & "Lista"
        Call NomeBrowse_Validate(bSGECancelDummy)
        
        Classe.Text = "Class" & NomeArq.Text
        
        NomeTela.Text = NomeArq.Text
        
        iAlteradoNomeArq = 0
    
    End If

    Exit Sub

Erro_NomeArq_Validate:

    Cancel = True

    Select Case gErr
    
        Case 131721 To 131722

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143923)

    End Select

    Exit Sub

End Sub

Private Sub NomeBrowse_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objBrowseArquivo As New AdmBrowseArquivo
Dim vbMsgResult As VbMsgBoxResult

On Error GoTo Erro_NomeBrowse_Validate
        
    gbCriarScriptDelete = False

    lErro = CF("BrowseArquivo_Le", NomeBrowse.Text, objBrowseArquivo)
    If lErro <> SUCESSO Then gError 131673
    
    If objBrowseArquivo.iBotaoEdita <> 0 Or objBrowseArquivo.iBotaoSeleciona <> 0 Or objBrowseArquivo.iBotaoConsulta <> 0 Then
        
        vbMsgResult = Rotina_Aviso(vbYesNo, "AVISO_BROWSE_JA_EXISTENTE")
        If vbMsgResult <> vbNo Then
            gbCriarScriptDelete = True
            Call Browse_Apaga
        End If
    End If

    Exit Sub

Erro_NomeBrowse_Validate:

    Cancel = True

    Select Case gErr
    
        Case 131673
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143924)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_Colunas(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Alocação

Dim iIndice As Integer

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Browse")
    objGridInt.colColuna.Add ("Indice")
    objGridInt.colColuna.Add ("Chave")
    objGridInt.colColuna.Add ("Coluna")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Tam")
    objGridInt.colColuna.Add ("Prec")
    objGridInt.colColuna.Add ("Ord")
    objGridInt.colColuna.Add ("Classe")
    objGridInt.colColuna.Add ("Atrib Classe")
    objGridInt.colColuna.Add ("Tam Tela")
    objGridInt.colColuna.Add ("SubTipo")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Browse.Name)
    objGridInt.colCampo.Add (Indice.Name)
    objGridInt.colCampo.Add (Chave.Name)
    objGridInt.colCampo.Add (Coluna.Name)
    objGridInt.colCampo.Add (Descricao.Name)
    objGridInt.colCampo.Add (Tipo.Name)
    objGridInt.colCampo.Add (Tamanho.Name)
    objGridInt.colCampo.Add (Precisao.Name)
    objGridInt.colCampo.Add (Ordem.Name)
    objGridInt.colCampo.Add (TemClasse.Name)
    objGridInt.colCampo.Add (AtribClasse.Name)
    objGridInt.colCampo.Add (TamanhoTela.Name)
    objGridInt.colCampo.Add (SubTipo.Name)
    
    'Colunas da Grid
    iGrid_Browse_Col = 1
    iGrid_Indice_Col = 2
    iGrid_Chave_Col = 3
    iGrid_Coluna_Col = 4
    iGrid_Descricao_Col = 5
    iGrid_Tipo_Col = 6
    iGrid_Tamanho_Col = 7
    iGrid_Precisao_Col = 8
    iGrid_Ordem_Col = 9
    iGrid_TemClasse_Col = 10
    iGrid_AtribClasse_Col = 11
    iGrid_TamanhoTela_Col = 12
    iGrid_SubTipo_Col = 13

    'Grid do GridInterno
    objGridInt.objGrid = GridColunas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridColunas.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Colunas = SUCESSO

    Exit Function

End Function

Public Function Trata_Parametros(Optional colColunasTabelas As Collection) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
     
    iAlterado = 0
    iAlteradoNomeArq = 0
    
    If Not (colColunasTabelas Is Nothing) Then
    
        lErro = Preenche_GridColunas(colColunasTabelas)
        If lErro <> SUCESSO Then gError 131742
    
    End If
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 131742
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143925)
            
    End Select
    
    iAlterado = 0
    iAlteradoNomeArq = 0
        
    Exit Function

End Function

Private Sub ObtemSiglaTipo(ByVal sTipoBD As String, sSigla As String, iTipoBrowse As Integer, sTipoVB As String)

    Select Case sTipoBD
    
        Case "datetime", "smalldatetime"
            sSigla = "dt"
            iTipoBrowse = 6
            sTipoVB = "Date"
        
        Case "float", "decimal", "money", "numeric", "real", "smallmoney"
            sSigla = "d"
            iTipoBrowse = 3
            sTipoVB = "Double"
        
        Case "int", "bigint"
            sSigla = "l"
            iTipoBrowse = 2
            sTipoVB = "Long"
        
        Case "smallint", "bit", "tinyint"
            sSigla = "i"
            iTipoBrowse = 1
            sTipoVB = "Integer"
                
        Case "varchar", "char", "nchar", "ntext", "nvarchar"
            sSigla = "s"
            iTipoBrowse = 4
            sTipoVB = "String"
            
        Case Else
            sSigla = "s"
            iTipoBrowse = 4
            sTipoVB = "String"
            
    End Select

End Sub

Public Function Move_Tela_Memoria(ByVal colColunasTabelas As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objColunasTabelas As ClassColunasTabelas

On Error GoTo Erro_Move_Tela_Memoria

    'Preenche colReservaItem com as linhas do GridReserva
    For iIndice = 1 To objGridColunas.iLinhasExistentes
    
        Set objColunasTabelas = New ClassColunasTabelas

        objColunasTabelas.iBrowse = StrParaInt(GridColunas.TextMatrix(iIndice, iGrid_Browse_Col))
        objColunasTabelas.sColuna = GridColunas.TextMatrix(iIndice, iGrid_Coluna_Col)
        objColunasTabelas.iOrdem = StrParaInt(GridColunas.TextMatrix(iIndice, iGrid_Ordem_Col))
        objColunasTabelas.lColunaPrecisao = StrParaLong(GridColunas.TextMatrix(iIndice, iGrid_Precisao_Col))
        objColunasTabelas.lColunaTamanho = StrParaLong(GridColunas.TextMatrix(iIndice, iGrid_Tamanho_Col))
        objColunasTabelas.sColunaTipo = GridColunas.TextMatrix(iIndice, iGrid_Tipo_Col)
        objColunasTabelas.sAtributoClasse = GridColunas.TextMatrix(iIndice, iGrid_AtribClasse_Col)
        objColunasTabelas.lTamanhoTela = StrParaLong(GridColunas.TextMatrix(iIndice, iGrid_TamanhoTela_Col))
        objColunasTabelas.sDescricao = GridColunas.TextMatrix(iIndice, iGrid_Descricao_Col)
        objColunasTabelas.iClasse = StrParaInt(GridColunas.TextMatrix(iIndice, iGrid_TemClasse_Col))
        objColunasTabelas.iIndice = StrParaInt(GridColunas.TextMatrix(iIndice, iGrid_Indice_Col))
        objColunasTabelas.iChave = StrParaInt(GridColunas.TextMatrix(iIndice, iGrid_Chave_Col))
        objColunasTabelas.iSubTipo = StrParaInt(GridColunas.TextMatrix(iIndice, iGrid_SubTipo_Col))
        
        colColunasTabelas.Add objColunasTabelas

    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143926)

    End Select

    Exit Function

End Function

Private Function MantemSequencialOrdem(ByVal iValorAntigo As Integer, ByVal iValorNovo As Integer, ByVal iLinha As Integer)

Dim iIndice As Integer
Dim iOrd As Integer

    If iValorAntigo <> iValorNovo Then
    
        'Se diminuiu o Valor
        If iValorAntigo > iValorNovo Then
        
            For iIndice = 1 To objGridColunas.iLinhasExistentes
    
                'Ignora linha alterada
                If iIndice <> iLinha Then
    
                    'Se o valor está entre o Valor Novo (inclusive) e o Antigo (exclusive)
                    If StrParaInt(GridColunas.TextMatrix(iIndice, iGrid_Ordem_Col)) >= iValorNovo And StrParaInt(GridColunas.TextMatrix(iIndice, iGrid_Ordem_Col)) < iValorAntigo Then
                    
                        'Incrementa
                        GridColunas.TextMatrix(iIndice, iGrid_Ordem_Col) = CStr(StrParaInt(GridColunas.TextMatrix(iIndice, iGrid_Ordem_Col)) + 1)
                    
                    End If
                    
                End If
                
            Next
        
        Else
        
            For iIndice = 1 To objGridColunas.iLinhasExistentes
    
                'Ignora linha alterada
                If iIndice <> iLinha Then
                
                    'Se o valor está entre o Valor Novo (inclusive) e o Antigo (exclusive)
                    If StrParaInt(GridColunas.TextMatrix(iIndice, iGrid_Ordem_Col)) <= iValorNovo And StrParaInt(GridColunas.TextMatrix(iIndice, iGrid_Ordem_Col)) > iValorAntigo Then
                    
                        'Decrementa
                        GridColunas.TextMatrix(iIndice, iGrid_Ordem_Col) = CStr(StrParaInt(GridColunas.TextMatrix(iIndice, iGrid_Ordem_Col)) - 1)
                    
                    End If
                    
                End If
                
            Next
        
        End If
    
    End If
    
End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Limpa a Tela
    Call Limpa_BrowseCria
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143927)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_BrowseCria()
'Limpa os campos da tela sem fechar o sistema de setas
    
    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridColunas)
    
    NomeArq.Text = ""
    ModuloTela.Text = ""
    ModuloClasse.Text = ""
    ModuloAcesso.Text = ""
    ModuloFormata.Text = ""
    
    ErroGerado.Caption = ""

    iAlterado = 0
    iAlteradoNomeArq = 0
    gbCriarScriptDelete = False
    
    gsTipoArquivo = ""
    iCont = 0
    gsErroLeitura = ""

    Set gobjTela = New ClassCriaTela

    Exit Sub

End Sub

'####################################################
'INICIO DOS GERADORES DE SCRIPTS
Private Function Browse_Apaga() As Long

Dim lErro As Long
Dim sNL As String
Dim sBrowseApaga As String

On Error GoTo Erro_Browse_Apaga

    sNL = Chr(10)

    sBrowseApaga = "--SCRIPT DE EXCLUSÃO DE BROWSE CRIADO AUTOMATICAMENTE PELA TELA BROWSE CRIA"
    sBrowseApaga = sBrowseApaga & sNL & "DELETE Arquivos " & sNL & "WHERE Nome = '" & NomeArq.Text & "'" & sNL & "GO"
    sBrowseApaga = sBrowseApaga & sNL & "DELETE Campos " & sNL & "WHERE NomeArq = '" & NomeArq.Text & "'" & sNL & "GO"
    sBrowseApaga = sBrowseApaga & sNL & "DELETE Telas " & sNL & "WHERE Nome = '" & NomeBrowse.Text & "'" & sNL & "GO"
    sBrowseApaga = sBrowseApaga & sNL & "DELETE GrupoTela " & sNL & "WHERE NomeTela = '" & NomeBrowse.Text & "'" & sNL & "GO"
    sBrowseApaga = sBrowseApaga & sNL & "DELETE TelasModulo " & sNL & "WHERE NomeTela = '" & NomeBrowse.Text & "'" & sNL & "GO"
    sBrowseApaga = sBrowseApaga & sNL & "DELETE BrowseArquivo " & sNL & "WHERE NomeTela = '" & NomeBrowse.Text & "'" & sNL & "GO"
    sBrowseApaga = sBrowseApaga & sNL & "DELETE BrowseCampo " & sNL & "WHERE NomeTela = '" & NomeBrowse.Text & "'" & sNL & "GO"
    sBrowseApaga = sBrowseApaga & sNL & "DELETE BrowseIndice " & sNL & "WHERE NomeTela = '" & NomeBrowse.Text & "'" & sNL & "GO"
    sBrowseApaga = sBrowseApaga & sNL & "DELETE BrowseUsuarioCampo " & sNL & "WHERE NomeTela = '" & NomeBrowse.Text & "'" & sNL & "GO"
    sBrowseApaga = sBrowseApaga & sNL & "DELETE GrupoBrowseCampo " & sNL & "WHERE NomeTela = '" & NomeBrowse.Text & "'" & sNL & "GO"

    ScriptBrowse.Text = sBrowseApaga

    Browse_Apaga = SUCESSO

    Exit Function

Erro_Browse_Apaga:

    Browse_Apaga = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143928)

    End Select
    
    Exit Function
    
End Function

Private Function Browse_Cria(ByVal colColunasTabelas As Collection, sBrowse As String) As Long

Dim lErro As Long
Dim objColunasTabelas As ClassColunasTabelas
Dim objColunasTabelasAux As ClassColunasTabelas
Dim sNL As String
Dim sSQL As String
Dim iIndice As Integer
Dim sSQLAux As String
Dim sTipoArq As String
Dim iSeq As Integer
Dim sSigla As String
Dim iTipo As Integer
Dim sTipoVB As String
Dim iPosTela As Integer
Dim bTemFilialEmpresa As Boolean
Dim sSelecao As String

On Error GoTo Erro_Browse_Cria

    sBrowse = "--BROWSE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA" & sNL
    
    sNL = Chr(10)
    
    sSQL = "INSERT INTO Campos(NomeArq,Nome,Descricao,Obrigatorio,Imexivel, Ativo, Tipo, Tamanho, Precisao, Decimais, TamExibicao, TituloEntradaDados, TituloGrid, Ordinal,Alinhamento, SubTipo ) " & sNL & "VALUES "
    iIndice = 0
    
    bTemFilialEmpresa = False
    sSelecao = "''"
    For Each objColunasTabelas In colColunasTabelas
        If UCase(objColunasTabelas.sColuna) = "FILIALEMPRESA" Then
            bTemFilialEmpresa = True
            sSelecao = "'FilialEmpresa = ?'"
            Exit For
        End If
    Next
    
    'ARQUIVO
    'TELAS
    'GRUPOTELAS
    'TELASMODULO
    'BROWSEARQUIVOS
    'CAMPOS
    For Each objColunasTabelas In colColunasTabelas
    
        iIndice = iIndice + 1
        
        With objColunasTabelas
    
            If iIndice = 1 Then
                
                If gsTipoArquivo = "U" Then
                    sTipoArq = "2" 'Tabela
                Else
                    sTipoArq = "4" 'View
                End If
            
                sBrowse = sBrowse & sNL & "--ARQUIVO "
                sBrowse = sBrowse & sNL & "INSERT INTO Arquivos (Nome, Descricao, Tipo) " & sNL & "VALUES ('" & NomeArq.Text & "','" & DescBrowse.Text & "'," & sTipoArq & ")" & sNL & "GO"
                
                sBrowse = sBrowse & sNL & sNL & "--TELAS "
                sBrowse = sBrowse & sNL & "INSERT INTO Telas (Nome, Projeto_Original, Classe_Original, FilialEmpresa, Descricao) " & sNL & "VALUES ('" & NomeBrowse.Text & "','Telas" & ModuloTela.Text & "','ClassTelas" & ModuloTela.Text & "',1,'" & DescBrowse.Text & "')" & sNL & "GO"
            
                sBrowse = sBrowse & sNL & sNL & "--GRUPOTELAS "
                sBrowse = sBrowse & sNL & "INSERT INTO GrupoTela (CodGrupo, NomeTela, TipoDeAcesso) " & sNL & "VALUES ('supervisor', '" & NomeBrowse.Text & "', 1)" & sNL & "GO"
            
                sBrowse = sBrowse & sNL & sNL & "--TELASMODULO"
                sBrowse = sBrowse & sNL & "INSERT INTO TelasModulo (SiglaModulo, NomeTela) " & sNL & "VALUES ('" & ModuloAcesso.Text & "', '" & NomeBrowse.Text & "')" & sNL & "GO"
            
                sBrowse = sBrowse & sNL & sNL & "--BROWSEARQUIVOS"
                sBrowse = sBrowse & sNL & "INSERT INTO BrowseArquivo (NomeTela, NomeArq, SelecaoSQL, Projeto, Classe, TituloBrowser, BotaoSeleciona, BotaoEdita, BotaoConsulta, ProjetoObjeto, ClasseObjeto, BancoDados, NomeTelaEdita, NomeTelaConsulta) " & sNL & "VALUES ('" & NomeBrowse.Text & "', '" & NomeArq.Text & "', " & sSelecao & ", 'Rotinas" & ModuloFormata.Text & "', 'Class" & ModuloFormata.Text & "Formata', 'Lista de " & DescBrowse.Text & "', " & CStr(Botao(0).Value) & ", " & CStr(Botao(1).Value) & ", " & CStr(Botao(2).Value) & ", 'Globais" & ModuloClasse.Text & "', '" & Classe.Text & "', 0, '" & NomeTela.Text & "', '" & NomeTelaConsulta.Text & "')" & sNL & "GO"
            
                If bTemFilialEmpresa Then
                    sBrowse = sBrowse & sNL & sNL & "--BROWSEPARAMSELECAO"
                    sBrowse = sBrowse & sNL & "INSERT INTO BrowseParamSelecao (NomeTela, Ordem, Projeto, Classe, Property) " & sNL & "VALUES ('" & NomeBrowse.Text & "',1,'admlib','Adm','giFilialEmpresa')" & sNL & "GO"
                End If
            
                sSQLAux = sNL & "--CAMPOS " & sNL
            Else
            
                sSQLAux = ""
            
            End If
            
            Call ObtemSiglaTipo(.sColunaTipo, sSigla, iTipo, sTipoVB)
            
            sSQLAux = sSQLAux & sSQL & "('" & NomeArq.Text & "', '" & .sColuna & "', '" & .sDescricao & "',1,0,1," & CStr(iTipo) & "," & CStr(.lColunaTamanho) & "," & CStr(.lColunaPrecisao) & "," & CStr(.lColunaPrecisao) & ",0, '" & .sDescricao & "', '" & .sDescricao & "', " & CStr(.iOrdem) & ",0," & CStr(.iSubTipo) & ")" & sNL & "GO"
            sBrowse = sBrowse & sNL & sSQLAux
        
        End With
    
    Next
    
    'BROWSECAMPO
    sSQL = "INSERT INTO BrowseCampo(NomeTela,NomeCampo,Nome)" & sNL & "VALUES "
    iIndice = 0
    sBrowse = sBrowse & sNL & sNL & "--BROWSECAMPO "
    
    For Each objColunasTabelas In colColunasTabelas
    
        iIndice = iIndice + 1
        
        With objColunasTabelas
        
            If .iClasse = MARCADO Then
            
                sSQLAux = sSQL & "('" & NomeBrowse.Text & "', '" & .sColuna & "', '" & .sAtributoClasse & "')" & sNL & "GO"
                sBrowse = sBrowse & sNL & sSQLAux
            
            End If
        
        End With
    
    Next
    
    'BROWSEUSUARIOCAMPO
    sSQL = "INSERT INTO BrowseUsuarioCampo (NomeTela, CodUsuario, NomeArq, Nome, PosicaoTela, Titulo, Largura)" & sNL & "VALUES "
    iIndice = 0
    sBrowse = sBrowse & sNL & sNL & "--BROWSEUSUARIOCAMPO "
    
    For Each objColunasTabelas In colColunasTabelas
    
        iIndice = iIndice + 1
                
        With objColunasTabelas
        
            If .iBrowse = MARCADO Then
            
                'Calcula a Posição na tela de acordo com a Ordem
                iPosTela = 1
                For Each objColunasTabelasAux In colColunasTabelas
                    If objColunasTabelasAux.iBrowse = MARCADO Then
                        If .iOrdem > objColunasTabelasAux.iOrdem Then iPosTela = iPosTela + 1
                    End If
                Next
            
                sSQLAux = sSQL & "('" & NomeBrowse.Text & "','supervisor', '" & NomeArq.Text & "', '" & .sColuna & "', " & iPosTela & ", '" & .sDescricao & "', " & CStr(.lTamanhoTela) & ")" & sNL & "GO"
                sBrowse = sBrowse & sNL & sSQLAux
        
            End If
        
        End With
    
    Next
    
    'GRUPOBROWSECAMPO
    sSQL = "INSERT INTO GrupoBrowseCampo (CodGrupo, NomeTela, NomeArq, Nome)" & sNL & "VALUES "
    iIndice = 0
    sBrowse = sBrowse & sNL & sNL & "--GRUPOBROWSECAMPO "
    
    For Each objColunasTabelas In colColunasTabelas
    
        iIndice = iIndice + 1
        
        With objColunasTabelas
        
            If Not (UCase(objColunasTabelas.sColuna) Like "NUMINT*") Then
            
                sSQLAux = sSQL & "('supervisor', '" & NomeBrowse.Text & "', '" & NomeArq.Text & "', '" & .sColuna & "')" & sNL & "GO"
                sBrowse = sBrowse & sNL & sSQLAux
            
            End If
        
        End With
    
    Next
    
    'BROWSEINDICE
    sSQL = "INSERT INTO BrowseIndice (NomeTela, Indice, NomeIndice, OrdenacaoSQL, SelecaoSQL) " & sNL & "VALUES "
    iIndice = 0
    sBrowse = sBrowse & sNL & sNL & "--BROWSEINDICE "
    iSeq = 0
    
    For Each objColunasTabelas In colColunasTabelas
    
        iIndice = iIndice + 1
        
        With objColunasTabelas
        
            If .iIndice = MARCADO Then
            
                iSeq = iSeq + 1
            
                sSQLAux = sSQL & "('" & NomeBrowse.Text & "'," & CStr(iSeq) & ", '" & .sDescricao & "', '" & .sColuna & "', '" & .sColuna & " < ?')" & sNL & "GO"
                sBrowse = sBrowse & sNL & sSQLAux
            
            End If
        
        End With
    
    Next
    
    Browse_Cria = SUCESSO

    Exit Function

Erro_Browse_Cria:

    Browse_Cria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143929)

    End Select
    
    Exit Function

End Function

Private Function ClasseType_Cria(ByVal colColunasTabelas As Collection, sClasse As String, sType As String) As Long

Dim lErro As Long
Dim objColunasTabelas As ClassColunasTabelas
Dim iIndice As Integer
Dim sNL As String
Dim sSigla As String
Dim iTipo As Integer
Dim sTipoVB As String

On Error GoTo Erro_ClasseType_Cria
    
    sNL = Chr(10)
    
    sType = "'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA" & sNL & "Type type" & Mid(Classe.Text, 6, Len(Classe.Text) - 5)
    sClasse = "'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA"
    
    For Each objColunasTabelas In colColunasTabelas
    
        iIndice = iIndice + 1
        
        With objColunasTabelas
        
            If .iClasse = MARCADO Then
            
                Call ObtemSiglaTipo(.sColunaTipo, sSigla, iTipo, sTipoVB)
                sClasse = sClasse & sNL & "Private mvar" & .sAtributoClasse & " AS " & sTipoVB
                sType = sType & sNL & TECLA_TAB & .sAtributoClasse & " AS " & sTipoVB
            
            End If
        
        End With
    
    Next

    sType = sType & sNL & "End Type"
    
    For Each objColunasTabelas In colColunasTabelas
    
        iIndice = iIndice + 1
        
        With objColunasTabelas
        
            If .iClasse = MARCADO Then
            
                Call ObtemSiglaTipo(.sColunaTipo, sSigla, iTipo, sTipoVB)
                sClasse = sClasse & sNL & sNL & "Public Property Let " & .sAtributoClasse & " (ByVal vData As " & sTipoVB & ")" & sNL & TECLA_TAB & "mvar" & .sAtributoClasse & " = vData" & sNL & "End Property"
                sClasse = sClasse & sNL & sNL & "Public Property Get " & .sAtributoClasse & " () AS " & sTipoVB & sNL & TECLA_TAB & .sAtributoClasse & "= mvar" & .sAtributoClasse & sNL & "End Property"

            End If
        
        End With
    
    Next
    
    ClasseType_Cria = SUCESSO

    Exit Function

Erro_ClasseType_Cria:

    ClasseType_Cria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143930)

    End Select
    
    Exit Function

End Function

Private Function RotinaLe_Cria(ByVal colColunasTabelas As Collection, sRotinaLe As String, sDicErros As String, sDicRotinas As String) As Long

Dim lErro As Long
Dim objColunasTabelas As ClassColunasTabelas
Dim iIndice As Integer
Dim sNL As String
Dim sSigla As String
Dim iTipo As Integer
Dim sTipoVB As String
Dim sSelect As String
Dim sWhere As String
Dim sWhereAux As String
Dim sSQLAux1 As String
Dim sSQLAux2 As String
Dim sMsg As String
Dim sMsgSet As String
Dim sMsgErro As String
Dim sOBJ As String
Dim sNomeFunc As String
Dim sType As String
Dim sProxErro As String
Dim iQtdChaves As Integer

On Error GoTo Erro_RotinaLe_Cria
    
    sOBJ = "obj" & Mid(Classe.Text, 6, Len(Classe.Text) - 5)
    sNomeFunc = NomeArq.Text & "_Le"
    sType = "t" & Mid(Classe.Text, 6, Len(Classe.Text) - 5)
    sNL = Chr(10)
    
    sDicRotinas = sDicRotinas & sNL & "INSERT INTO ROTINAS (Sigla, projeto_original, classe_original) " & sNL & "VALUES ('" & sNomeFunc & "','Rotinas" & ModuloClasse.Text & "','Class" & ModuloClasse.Text & "Select')" & sNL & "GO"
    sDicRotinas = sDicRotinas & sNL & "INSERT INTO RotinasModulo (SiglaModulo, SiglaRotina) " & sNL & "VALUES ('" & ModuloAcesso.Text & "', '" & sNomeFunc & "')" & sNL & "GO"
    sDicRotinas = sDicRotinas & sNL & "INSERT INTO GrupoRotinas (CodGrupo, SiglaRotina, TipoDeAcesso) " & sNL & "VALUES ('supervisor','" & sNomeFunc & "',1)" & sNL & "GO" & sNL
    
    sMsgErro = TECLA_TAB & "Select Case gerr" & sNL
    sRotinaLe = sRotinaLe & sNL & "'LEITURA"
    
    'Função
    sRotinaLe = sRotinaLe & sNL & "Public Function " & sNomeFunc & "(ByVal " & sOBJ & " As " & Classe.Text & ") As Long" & sNL
    
    'Declaração
    sRotinaLe = sRotinaLe & sNL & "Dim lErro As Long" & sNL & "Dim lComando As Long" & sNL & "Dim " & sType & " As type" & Mid(Classe.Text, 6, Len(Classe.Text) - 5) & sNL

    'On Error
    sRotinaLe = sRotinaLe & sNL & "On Error GoTo Erro_" & sNomeFunc & sNL

    'Abertura de Comando
    Call CalculaProximoErro(sProxErro)
    sRotinaLe = sRotinaLe & sNL & TECLA_TAB & "'Executa a abertura do Comando" & sNL & TECLA_TAB & "lComando = Comando_Abrir()" & sNL & TECLA_TAB & "If lComando = 0 Then gError " & sProxErro & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, """ & "ERRO_ABERTURA_COMANDO" & """, gErr)" & sNL

    'Alocação de espaço no buffer
    sMsg = TECLA_TAB & "'Alocação de espaço no buffer"
    sSelect = "SELECT "
    sSQLAux1 = ""
    sSQLAux2 = ""
    sWhere = "WHERE "
    iQtdChaves = 0
    sMsgSet = ""
    
    For Each objColunasTabelas In colColunasTabelas
    
        iIndice = iIndice + 1
        
        With objColunasTabelas
        
            Call ObtemSiglaTipo(.sColunaTipo, sSigla, iTipo, sTipoVB)
            If sSigla = "s" Then
                sMsg = sMsg & sNL & TECLA_TAB & sType & "." & .sAtributoClasse & " = String(UTILIZAR_STRING_TAMANHO_" & CStr(.lColunaTamanho) & ",0)"
            End If
        
            If iIndice <> 1 Then
                sSelect = sSelect & ", "
                sSQLAux1 = sSQLAux1 & ", "
            End If
            
            'Quebra linha  a cada 7 colunas
            If iIndice Mod 7 = 0 Then
                sSelect = sSelect & """" & " & _" & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & TECLA_TAB & """"
            End If
            
            'Quebra a linha a cada 5 types
            If iIndice Mod 5 = 0 Then
                sSQLAux1 = sSQLAux1 & " _ " & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & TECLA_TAB
            End If
            
            sSelect = sSelect & .sColuna
            sSQLAux1 = sSQLAux1 & sType & "." & .sAtributoClasse
            sMsgSet = sMsgSet & sNL & TECLA_TAB & sOBJ & "." & .sAtributoClasse & " = " & sType & "." & .sAtributoClasse
        
            If .iChave = MARCADO Then
            
                iQtdChaves = iQtdChaves + 1
            
                If iQtdChaves <> 1 Then
                    sSQLAux2 = sSQLAux2 & ", "
                    sWhere = sWhere & " AND "
                    sWhereAux = sWhereAux & " e "
                End If
            
                sSQLAux2 = sSQLAux2 & sOBJ & "." & .sAtributoClasse
                sWhere = sWhere & .sColuna & "= ? "
                sWhereAux = sWhereAux & .sColuna & " %s"
           
            End If
        
        End With
    
    Next

    'Aloca espaço no buffer
    sRotinaLe = sRotinaLe & sNL & sMsg & sNL
    
    'Select
    Call CalculaProximoErro(sProxErro)
    sRotinaLe = sRotinaLe & sNL & TECLA_TAB & "'Le a tabela" & NomeArq.Text & sNL & TECLA_TAB & "lErro = Comando_Executar(lComando, """ & sSelect & " FROM " & NomeArq.Text & " " & sWhere & """" & ", _ " & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & TECLA_TAB & sSQLAux1 & ", _" & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & TECLA_TAB & sSQLAux2 & ")"
    sRotinaLe = sRotinaLe & sNL & TECLA_TAB & "If lErro <> AD_SQL_SUCESSO Then gError " & sProxErro & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro
    
    'Busca Primeiro
    Call CalculaProximoErro(sProxErro)
    sRotinaLe = sRotinaLe & sNL & TECLA_TAB & "'Busca Primeiro" & sNL & TECLA_TAB & "lErro = Comando_BuscarPrimeiro(lComando)"
    sRotinaLe = sRotinaLe & sNL & TECLA_TAB & "If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError " & sProxErro & sNL
    sMsgErro = sMsgErro & ", " & sProxErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, """ & "ERRO_LEITURA_" & UCase(NomeArq.Text) & """" & ", gErr)"
    sDicErros = sDicErros & sNL & "INSERT INTO Erros (Codigo, Descricao) " & sNL & "VALUES ('ERRO_LEITURA_" & UCase(NomeArq.Text) & "', 'Não foi possível ler a Tabela " & DescBrowse.Text & "')" & sNL & "GO"
    
    'Sem dados
    Call CalculaProximoErro(sProxErro)
    'gsErroLeitura = sProxErro
    sRotinaLe = sRotinaLe & sNL & TECLA_TAB & "'Sem Dados" & sNL & TECLA_TAB & "If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS" '& sProxErro
    sMsgErro = sMsgErro & sNL & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro & " 'Sem dados -> Tratado na rotina chamadora" & sNL
    sDicErros = sDicErros & sNL & "INSERT INTO Erros (Codigo, Descricao) " & sNL & "VALUES ('ERRO_" & UCase(NomeArq.Text) & "_NAO_CADASTRADO', 'Não foi possível ler o(a) " & DescBrowse.Text & " de " & sWhereAux & ".')" & sNL & "GO"
    
    sRotinaLe = sRotinaLe & sNL & sMsgSet & sNL
    
    'Fecha comando
    sRotinaLe = sRotinaLe & sNL & TECLA_TAB & "'Fecha Comando" & sNL & TECLA_TAB & "Call Comando_Fechar(lComando)" & sNL
        
    'SUCESSO
    sRotinaLe = sRotinaLe & sNL & TECLA_TAB & sNomeFunc & " = SUCESSO" & sNL
        
    'Exit Function
    sRotinaLe = sRotinaLe & sNL & TECLA_TAB & "Exit Function" & sNL
        
    'Label
    sRotinaLe = sRotinaLe & sNL & "Erro_" & sNomeFunc & ":" & sNL
        
    'Erro
    sRotinaLe = sRotinaLe & sNL & TECLA_TAB & sNomeFunc & " = gerr" & sNL
   
    Call CalculaProximoErro(sProxErro)
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case Else"
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, " & """" & "ERRO_FORNECIDO_PELO_VB" & """" & ", gErr, Error, " & sProxErro & ")" & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & "End Select"

    'Select Case gerr
    sRotinaLe = sRotinaLe & sNL & sMsgErro & sNL

    'Fecha comando
    sRotinaLe = sRotinaLe & sNL & TECLA_TAB & "'Fecha Comando" & sNL & TECLA_TAB & "Call Comando_Fechar(lComando)" & sNL

    'Exit Function
    sRotinaLe = sRotinaLe & sNL & TECLA_TAB & "Exit Function" & sNL
    
    'End Function
    sRotinaLe = sRotinaLe & sNL & "End Function"
    
    RotinaLe_Cria = SUCESSO

    Exit Function

Erro_RotinaLe_Cria:

    RotinaLe_Cria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143932)

    End Select
    
    Exit Function

End Function

Private Function RotinaExclui_Cria(ByVal colColunasTabelas As Collection, sRotinaExclui As String, sDicErros As String, sDicRotinas As String) As Long

Dim lErro As Long
Dim objColunasTabelas As ClassColunasTabelas
Dim iIndice As Integer
Dim sNL As String
Dim sSigla As String
Dim iTipo As Integer
Dim sTipoVB As String
Dim sSelect As String
Dim sWhere As String
Dim sFiltro As String
Dim sOBJ As String
Dim sNomeFunc As String
Dim sType As String
Dim sProxErro As String
Dim iQtdChaves As Integer
Dim sErroCampos As String
Dim sMsgErro As String
Dim sChave As String

On Error GoTo Erro_RotinaExclui_Cria

    sOBJ = "obj" & Mid(Classe.Text, 6, Len(Classe.Text) - 5)
    sType = "t" & Mid(Classe.Text, 6, Len(Classe.Text) - 5)
    sNL = Chr(10)
    
    sMsgErro = TECLA_TAB & "Select Case gerr" & sNL
    sRotinaExclui = sRotinaExclui & sNL & "'EXCLUSÃO"
    
    '#############################################################
    'FUNÇÃO QUE CHAMA A EXCLUSÃO EM TRANSAÇÃO
    sNomeFunc = NomeArq.Text & "_Exclui"
    
    sDicRotinas = sDicRotinas & sNL & "INSERT INTO ROTINAS (Sigla, projeto_original, classe_original) " & sNL & "VALUES ('" & sNomeFunc & "','Rotinas" & ModuloClasse.Text & "','Class" & ModuloClasse.Text & "Grava')" & sNL & "GO"
    sDicRotinas = sDicRotinas & sNL & "INSERT INTO RotinasModulo (SiglaModulo, SiglaRotina) " & sNL & "VALUES ('" & ModuloAcesso.Text & "', '" & sNomeFunc & "')" & sNL & "GO"
    sDicRotinas = sDicRotinas & sNL & "INSERT INTO GrupoRotinas (CodGrupo, SiglaRotina, TipoDeAcesso) " & sNL & "VALUES ('supervisor','" & sNomeFunc & "',1)" & sNL & "GO" & sNL
    
    'Função
    sRotinaExclui = sRotinaExclui & sNL & "Public Function " & sNomeFunc & "(ByVal " & sOBJ & " As " & Classe.Text & ") As Long" & sNL
    
    'Declaração
    sRotinaExclui = sRotinaExclui & sNL & "Dim lErro As Long" & sNL & "Dim lTransacao As Long" & sNL

    'On Error
    sRotinaExclui = sRotinaExclui & sNL & "On Error GoTo Erro_" & sNomeFunc & sNL

    'Abertura de Transação
    Call CalculaProximoErro(sProxErro)
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "'Abertura de transação"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "lTransacao = Transacao_Abrir()"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "If lTransacao = 0 Then gError " & sProxErro & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, """ & "ERRO_ABERTURA_TRANSACAO" & """, gErr)" & sNL

    'Chama gravação em transação
    Call CalculaProximoErro(sProxErro)
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "lErro = CF(" & """" & sNomeFunc & "_EmTrans" & """" & ", " & sOBJ & ")"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "If lErro <> SUCESSO Then gError " & sProxErro & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro & sNL

    'Confirma a transação
    Call CalculaProximoErro(sProxErro)
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "'Confirma a transação"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "lErro = Transacao_Commit()"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "If lErro <> AD_SQL_SUCESSO Then gError " & sProxErro & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, """ & "ERRO_COMMIT" & """, gErr)" & sNL

    'Sucesso
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & sNomeFunc & " = SUCESSO" & sNL
    
    'Exit Function
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "Exit Function" & sNL
        
    'Label
    sRotinaExclui = sRotinaExclui & sNL & "Erro_" & sNomeFunc & ":" & sNL
        
    'Erro
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & sNomeFunc & " = gerr" & sNL
        
    Call CalculaProximoErro(sProxErro)
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case Else"
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, " & """" & "ERRO_FORNECIDO_PELO_VB" & """" & ", gErr, Error, " & sProxErro & ")" & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & "End Select"

    'Select Case gerr
    sRotinaExclui = sRotinaExclui & sNL & sMsgErro & sNL

    'Fecha transação
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "'Desfaz Transação" & sNL & TECLA_TAB & "Call Transacao_Rollback" & sNL

    'Exit Function
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "Exit Function" & sNL
    
    'End Function
    sRotinaExclui = sRotinaExclui & sNL & "End Function" & sNL
    
    '################################################################
    'FUNÇÃO EM TRANSAÇÃO
    sNomeFunc = sNomeFunc & "_EmTrans"
       
    sDicRotinas = sDicRotinas & sNL & "INSERT INTO ROTINAS (Sigla, projeto_original, classe_original) " & sNL & "VALUES ('" & sNomeFunc & "','Rotinas" & ModuloClasse.Text & "','Class" & ModuloClasse.Text & "Grava')" & sNL & "GO"
    sDicRotinas = sDicRotinas & sNL & "INSERT INTO RotinasModulo (SiglaModulo, SiglaRotina) " & sNL & "VALUES ('" & ModuloAcesso.Text & "', '" & sNomeFunc & "')" & sNL & "GO"
    sDicRotinas = sDicRotinas & sNL & "INSERT INTO GrupoRotinas (CodGrupo, SiglaRotina, TipoDeAcesso) " & sNL & "VALUES ('supervisor','" & sNomeFunc & "',1)" & sNL & "GO" & sNL
    
    sMsgErro = TECLA_TAB & "Select Case gerr" & sNL
   
    'Função
    sRotinaExclui = sRotinaExclui & sNL & "Public Function " & sNomeFunc & "(ByVal " & sOBJ & " As " & Classe.Text & ") As Long" & sNL
    
    'Declaração
    sRotinaExclui = sRotinaExclui & sNL & "Dim lErro As Long" & sNL & "Dim alComando(0 To 1) As Long" & sNL & "Dim iIndice As Integer" & sNL & "Dim iAux As Integer" & sNL

    'On Error
    sRotinaExclui = sRotinaExclui & sNL & "On Error GoTo Erro_" & sNomeFunc & sNL
    
    'Abertura de Comando
    Call CalculaProximoErro(sProxErro)
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "'Abertura de Comando"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "For iIndice = LBound(alComando) To UBound(alComando)"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & TECLA_TAB & "alComando(iIndice) = Comando_Abrir()"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & TECLA_TAB & "If alComando(iIndice) = 0 Then gError " & sProxErro
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "Next" & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, """ & "ERRO_ABERTURA_COMANDO" & """, gErr)" & sNL

    sFiltro = ""
    sWhere = "WHERE "
    iQtdChaves = 0
    iIndice = 0
    sChave = ""
    
    For Each objColunasTabelas In colColunasTabelas
        
        iIndice = iIndice + 1
        
        With objColunasTabelas
            
            If .iChave = MARCADO Then
            
                iQtdChaves = iQtdChaves + 1
            
                If iQtdChaves <> 1 Then
                    sFiltro = sFiltro & ", "
                    sWhere = sWhere & " AND "
                    sErroCampos = sErroCampos & " e "
                    sChave = ", "
                End If
            
                sFiltro = sFiltro & sOBJ & "." & .sAtributoClasse
                sWhere = sWhere & .sColuna & "= ? "
                sErroCampos = sErroCampos & .sColuna & " %s"
             
            End If
            
        End With
    
    Next

    'Select
    Call CalculaProximoErro(sProxErro)
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "'Le a tabela" & NomeArq.Text
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "lErro = Comando_ExecutarPos(alComando(0), """ & "SELECT 1" & " FROM " & NomeArq.Text & " " & sWhere & """" & ", _ " & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & TECLA_TAB & "0, iAux, " & sFiltro & ")"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "If lErro <> AD_SQL_SUCESSO Then gError " & sProxErro & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro
    
    'Busca Primeiro
    Call CalculaProximoErro(sProxErro)
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "'Busca Primeiro"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "lErro = Comando_BuscarPrimeiro(alComando(0))"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError " & sProxErro & sNL
    sMsgErro = sMsgErro & ", " & sProxErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, """ & "ERRO_LEITURA_" & UCase(NomeArq.Text) & """" & ", gErr)" & sNL
    
    'Se não existir = > Erro
    Call CalculaProximoErro(sProxErro)
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "'Se não existir => ERRO"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "If lErro = AD_SQL_SEM_DADOS Then gError " & sProxErro & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, """ & "ERRO_" & UCase(NomeArq.Text) & "_NAO_CADASTRADO" & """" & ", gErr, " & sFiltro & ")" & sNL
    
    'Lock
    Call CalculaProximoErro(sProxErro)
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "'Trava registro contra alterações/Leituras"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "lErro = Comando_LockExclusive(alComando(0))"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "If lErro <> SUCESSO Then gError " & sProxErro & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, """ & "ERRO_LOCKEXCLUSIVE_" & UCase(NomeArq.Text) & """" & ", gErr)" & sNL
    sDicErros = sDicErros & sNL & "INSERT INTO Erros (Codigo, Descricao) " & sNL & "VALUES ('ERRO_LOCKEXCLUSIVE_" & UCase(NomeArq.Text) & "', 'Erro ao fazer Lock na Tabela " & DescBrowse.Text & "')" & sNL & "GO"
    
    'Delete
    Call CalculaProximoErro(sProxErro)
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "lErro = Comando_ExecutarPos(alComando(1), " & """" & "DELETE FROM " & NomeArq.Text & """" & ", alComando(0))"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "If lErro <> AD_SQL_SUCESSO Then gError " & sProxErro & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, """ & "ERRO_EXCLUSAO_" & UCase(NomeArq.Text) & """" & ", gErr)" & sNL
    sDicErros = sDicErros & sNL & "INSERT INTO Erros (Codigo, Descricao) " & sNL & "VALUES ('ERRO_EXCLUSAO_" & UCase(NomeArq.Text) & "', 'Não foi possível excluir o registo na Tabela " & DescBrowse.Text & "')" & sNL & "GO"
    
    'Fecha comando
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "'Fecha Comando"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "For iIndice = LBound(alComando) To UBound(alComando)"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & TECLA_TAB & "Call Comando_Fechar(alComando(iIndice))"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "Next" & sNL
        
    'SUCESSO
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & sNomeFunc & " = SUCESSO" & sNL
        
    'Exit Function
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "Exit Function" & sNL
        
    'Label
    sRotinaExclui = sRotinaExclui & sNL & "Erro_" & sNomeFunc & ":" & sNL
        
    'Erro
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & sNomeFunc & " = gerr" & sNL
        
    Call CalculaProximoErro(sProxErro)
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case Else"
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, " & """" & "ERRO_FORNECIDO_PELO_VB" & """" & ", gErr, Error, " & sProxErro & ")" & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & "End Select"

    'Select Case gerr
    sRotinaExclui = sRotinaExclui & sNL & sMsgErro & sNL

    'Fecha comando
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "'Fecha Comando"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "For iIndice = LBound(alComando) To UBound(alComando)"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & TECLA_TAB & "Call Comando_Fechar(alComando(iIndice))"
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "Next" & sNL

    'Exit Function
    sRotinaExclui = sRotinaExclui & sNL & TECLA_TAB & "Exit Function" & sNL
    
    'End Function
    sRotinaExclui = sRotinaExclui & sNL & "End Function"
    
    RotinaExclui_Cria = SUCESSO

    Exit Function

Erro_RotinaExclui_Cria:

    RotinaExclui_Cria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143935)

    End Select
    
    Exit Function

End Function

Private Function RotinaGrava_Cria(ByVal colColunasTabelas As Collection, sRotinaGrava As String, sDicErros As String, sDicRotinas As String) As Long

Dim lErro As Long
Dim objColunasTabelas As ClassColunasTabelas
Dim iIndice As Integer
Dim sNL As String
Dim sSigla As String
Dim iTipo As Integer
Dim sTipoVB As String
Dim sSelect As String
Dim sWhere As String
Dim sFiltro As String
Dim sOBJ As String
Dim sNomeFunc As String
Dim sType As String
Dim sProxErro As String
Dim iQtdChaves As Integer
Dim iQtdColunas As Integer
Dim sInsertCampos As String
Dim sUpdate As String
Dim sValoresInsert As String
Dim sValoresUpdate As String
Dim sErroCampos As String
Dim sInterroga As String
Dim sMsgErro As String
Dim sObterNumIntDoc As String
Dim bTemNumIntDoc As Boolean

On Error GoTo Erro_RotinaGrava_Cria

    sOBJ = "obj" & Mid(Classe.Text, 6, Len(Classe.Text) - 5)
    sType = "t" & Mid(Classe.Text, 6, Len(Classe.Text) - 5)
    sNL = Chr(10)
    
    sMsgErro = TECLA_TAB & "Select Case gerr" & sNL
    sRotinaGrava = sRotinaGrava & sNL & "'GRAVAÇÃO"
    
    '#############################################################
    'FUNÇÃO QUE CHAMA A GRAVAÇÃO EM TRANSAÇÃO
    sNomeFunc = NomeArq.Text & "_Grava"
    
    sDicRotinas = sDicRotinas & sNL & "INSERT INTO ROTINAS (Sigla, projeto_original, classe_original) " & sNL & "VALUES ('" & sNomeFunc & "','Rotinas" & ModuloClasse.Text & "','Class" & ModuloClasse.Text & "Grava')" & sNL & "GO"
    sDicRotinas = sDicRotinas & sNL & "INSERT INTO RotinasModulo (SiglaModulo, SiglaRotina) " & sNL & "VALUES ('" & ModuloAcesso.Text & "', '" & sNomeFunc & "')" & sNL & "GO"
    sDicRotinas = sDicRotinas & sNL & "INSERT INTO GrupoRotinas (CodGrupo, SiglaRotina, TipoDeAcesso) " & sNL & "VALUES ('supervisor','" & sNomeFunc & "',1)" & sNL & "GO" & sNL
    
    'Função
    sRotinaGrava = sRotinaGrava & sNL & "Public Function " & sNomeFunc & "(ByVal " & sOBJ & " As " & Classe.Text & ") As Long" & sNL
    
    'Declaração
    sRotinaGrava = sRotinaGrava & sNL & "Dim lErro As Long" & sNL & "Dim lTransacao As Long" & sNL

    'On Error
    sRotinaGrava = sRotinaGrava & sNL & "On Error GoTo Erro_" & sNomeFunc & sNL

    'Abertura de Transação
    Call CalculaProximoErro(sProxErro)
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "'Abertura de transação"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "lTransacao = Transacao_Abrir()"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "If lTransacao = 0 Then gError " & sProxErro & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, """ & "ERRO_ABERTURA_TRANSACAO" & """, gErr)" & sNL

    'Chama gravação em transação
    Call CalculaProximoErro(sProxErro)
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "lErro = CF(" & """" & sNomeFunc & "_EmTrans" & """" & ", " & sOBJ & ")"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "If lErro <> SUCESSO Then gError " & sProxErro & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro & sNL

    'Confirma a transação
    Call CalculaProximoErro(sProxErro)
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "'Confirma a transação"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "lErro = Transacao_Commit()"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "If lErro <> AD_SQL_SUCESSO Then gError " & sProxErro & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, """ & "ERRO_COMMIT" & """, gErr)" & sNL

    'Sucesso
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & sNomeFunc & " = SUCESSO" & sNL
    
    'Exit Function
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "Exit Function" & sNL
        
    'Label
    sRotinaGrava = sRotinaGrava & sNL & "Erro_" & sNomeFunc & ":" & sNL
        
    'Erro
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & sNomeFunc & " = gerr" & sNL
        
    Call CalculaProximoErro(sProxErro)
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case Else"
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, " & """" & "ERRO_FORNECIDO_PELO_VB" & """" & ", gErr, Error, " & sProxErro & ")" & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & "End Select"

    'Select Case gerr
    sRotinaGrava = sRotinaGrava & sNL & sMsgErro & sNL

    'Fecha comando
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "'Desfaz Transação" & sNL & TECLA_TAB & "Call Transacao_Rollback" & sNL

    'Exit Function
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "Exit Function" & sNL
    
    'End Function
    sRotinaGrava = sRotinaGrava & sNL & "End Function" & sNL
    
    '################################################################
    'FUNÇÃO EM TRANSAÇÃO
    sNomeFunc = sNomeFunc & "_EmTrans"
    
    sDicRotinas = sDicRotinas & sNL & "INSERT INTO ROTINAS (Sigla, projeto_original, classe_original) " & sNL & "VALUES ('" & sNomeFunc & "','Rotinas" & ModuloClasse.Text & "','Class" & ModuloClasse.Text & "Grava')" & sNL & "GO"
    sDicRotinas = sDicRotinas & sNL & "INSERT INTO RotinasModulo (SiglaModulo, SiglaRotina) " & sNL & "VALUES ('" & ModuloAcesso.Text & "', '" & sNomeFunc & "')" & sNL & "GO"
    sDicRotinas = sDicRotinas & sNL & "INSERT INTO GrupoRotinas (CodGrupo, SiglaRotina, TipoDeAcesso) " & sNL & "VALUES ('supervisor','" & sNomeFunc & "',1)" & sNL & "GO" & sNL
    
    sMsgErro = TECLA_TAB & "Select Case gerr" & sNL
   
    'Função
    sRotinaGrava = sRotinaGrava & sNL & "Public Function " & sNomeFunc & "(ByVal " & sOBJ & " As " & Classe.Text & ") As Long" & sNL
    
    iIndice = 0
    sObterNumIntDoc = ""
    bTemNumIntDoc = False
    
    For Each objColunasTabelas In colColunasTabelas
        
        iIndice = iIndice + 1
        
        If UCase(objColunasTabelas.sColuna) = "NUMINTDOC" Then
    
            Call CalculaProximoErro(sProxErro)
            sObterNumIntDoc = TECLA_TAB & TECLA_TAB & "'Obter NumIntDoc"
            sObterNumIntDoc = sObterNumIntDoc & sNL & TECLA_TAB & TECLA_TAB & "lErro = CF(" & """" & "Config_ObterNumInt" & """" & ", " & """" & ModuloClasse & "Config" & """" & ", <" & """" & "NUM_INT_PROX_" & UCase(NomeArq.Text) & """" & ">, lNumIntDoc)"
            sObterNumIntDoc = sObterNumIntDoc & sNL & TECLA_TAB & TECLA_TAB & "If lErro <> SUCESSO Then gError " & sProxErro & sNL
            sObterNumIntDoc = sObterNumIntDoc & sNL & TECLA_TAB & TECLA_TAB & sOBJ & "." & objColunasTabelas.sAtributoClasse & " = lNumIntDoc" & sNL
            sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro & sNL
            bTemNumIntDoc = True
    
        End If
    
    Next

    'Declaração
    sRotinaGrava = sRotinaGrava & sNL & "Dim lErro As Long" & sNL & "Dim alComando(0 To 1) As Long" & sNL & "Dim iIndice As Integer" & sNL & "Dim iAux As Integer" & sNL

    If bTemNumIntDoc Then
        sRotinaGrava = sRotinaGrava & "Dim lNumIntDoc As Long" & sNL
    End If

    'On Error
    sRotinaGrava = sRotinaGrava & sNL & "On Error GoTo Erro_" & sNomeFunc & sNL
    
    'Abertura de Comando
    Call CalculaProximoErro(sProxErro)
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "'Abertura de Comando"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "For iIndice = LBound(alComando) To UBound(alComando)"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & TECLA_TAB & "alComando(iIndice) = Comando_Abrir()"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & TECLA_TAB & "If alComando(iIndice) = 0 Then gError " & sProxErro
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "Next" & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, """ & "ERRO_ABERTURA_COMANDO" & """, gErr)" & sNL

    sFiltro = ""
    sWhere = "WHERE "
    iQtdChaves = 0
    iQtdColunas = 0
    sInterroga = ""
    sInsertCampos = ""
    sUpdate = ""
    iIndice = 0
    
    For Each objColunasTabelas In colColunasTabelas
        
        iIndice = iIndice + 1
        
        With objColunasTabelas
            
            If .iChave = MARCADO Then
            
                iQtdChaves = iQtdChaves + 1
            
                If iQtdChaves <> 1 Then
                    sFiltro = sFiltro & ", "
                    sWhere = sWhere & " AND "
                    sErroCampos = sErroCampos & " e "
                End If
            
                sFiltro = sFiltro & sOBJ & "." & .sAtributoClasse
                sWhere = sWhere & .sColuna & "= ? "
                sErroCampos = sErroCampos & .sColuna & " %s"
                
            Else
                            
                If UCase(.sColuna) <> "NUMINTDOC" Then
    
                    iQtdColunas = iQtdColunas + 1
    
                    If iQtdColunas <> 1 Then
                        sUpdate = sUpdate & ", "
                        sValoresUpdate = sValoresUpdate & ", "
                    End If
                    
                    If iQtdColunas Mod 5 = 0 Then
                        sUpdate = sUpdate & """" & " & _" & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & TECLA_TAB & """"
                        sValoresUpdate = sValoresUpdate & " _ " & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & TECLA_TAB
                    End If
                    
                    sValoresUpdate = sValoresUpdate & sOBJ & "." & .sAtributoClasse
                    sUpdate = sUpdate & .sColuna & "= ? "
                    
                End If
            
            End If
            
             If iIndice <> 1 Then
                sInsertCampos = sInsertCampos & ", "
                sValoresInsert = sValoresInsert & ", "
                sInterroga = sInterroga & ","
            End If
            
            If iIndice Mod 5 = 0 Then
                sInsertCampos = sInsertCampos & """" & " & _" & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & TECLA_TAB & """"
                sValoresInsert = sValoresInsert & " _ " & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & TECLA_TAB
            End If
            
            sInsertCampos = sInsertCampos & .sColuna
            sValoresInsert = sValoresInsert & sOBJ & "." & .sAtributoClasse
            sInterroga = sInterroga & "?"
           
        End With
    
    Next

    'Select
    Call CalculaProximoErro(sProxErro)
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "'Le a tabela" & NomeArq.Text
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "lErro = Comando_ExecutarPos(alComando(0), """ & "SELECT 1" & " FROM " & NomeArq.Text & " " & sWhere & """" & ", _ " & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & TECLA_TAB & "0, iAux, " & sFiltro & ")"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "If lErro <> AD_SQL_SUCESSO Then gError " & sProxErro & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro
    
    'Busca Primeiro
    Call CalculaProximoErro(sProxErro)
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "'Busca Primeiro"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "lErro = Comando_BuscarPrimeiro(alComando(0))"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError " & sProxErro & sNL
    sMsgErro = sMsgErro & ", " & sProxErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, """ & "ERRO_LEITURA_" & UCase(NomeArq.Text) & """" & ", gErr)" & sNL
    
    'Se existir
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "'Se existir => UPDATE, senão => INSERT"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "If lErro = AD_SQL_SUCESSO Then " & sNL
    
    'Lock
    Call CalculaProximoErro(sProxErro)
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & TECLA_TAB & "'Trava registro contra alterações/Leituras"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & TECLA_TAB & "lErro = Comando_LockExclusive(alComando(0))"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & TECLA_TAB & "If lErro <> SUCESSO Then gError " & sProxErro & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, """ & "ERRO_LOCKEXCLUSIVE_" & UCase(NomeArq.Text) & """" & ", gErr)" & sNL
    
    'Update
    Call CalculaProximoErro(sProxErro)
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & TECLA_TAB & "lErro = Comando_ExecutarPos(alComando(1), " & """" & "UPDATE " & NomeArq.Text & " SET " & sUpdate & """" & ", alComando(0),  _" & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & TECLA_TAB & sValoresUpdate & ")"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & TECLA_TAB & "If lErro <> AD_SQL_SUCESSO Then gError " & sProxErro & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, """ & "ERRO_UPDATE_" & UCase(NomeArq.Text) & """" & ", gErr)" & sNL
    sDicErros = sDicErros & sNL & "INSERT INTO Erros (Codigo, Descricao) " & sNL & "VALUES ('ERRO_UPDATE_" & UCase(NomeArq.Text) & "', 'Não foi possível atualizar a Tabela " & DescBrowse.Text & "')" & sNL & "GO"
    
    'Senão
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "Else " & sNL
    
    'Insert
    sRotinaGrava = sRotinaGrava & sNL & sObterNumIntDoc
    
    Call CalculaProximoErro(sProxErro)
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & TECLA_TAB & "lErro = Comando_Executar(alComando(1), " & """" & "INSERT INTO " & NomeArq.Text & "( " & sInsertCampos & ") VALUES (" & sInterroga & ")" & """" & ", _" & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & TECLA_TAB & sValoresInsert & ")"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & TECLA_TAB & "If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError " & sProxErro & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, """ & "ERRO_INSERCAO_" & UCase(NomeArq.Text) & """" & ", gErr)" & sNL
    sDicErros = sDicErros & sNL & "INSERT INTO Erros (Codigo, Descricao) " & sNL & "VALUES ('ERRO_INSERCAO_" & UCase(NomeArq.Text) & "', 'Não foi possível inserir na Tabela " & DescBrowse.Text & "')" & sNL & "GO"
    
    'Fim Se
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "End If " & sNL
    
    'Fecha comando
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "'Fecha Comando"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "For iIndice = LBound(alComando) To UBound(alComando)"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & TECLA_TAB & "Call Comando_Fechar(alComando(iIndice))"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "Next" & sNL
        
    'SUCESSO
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & sNomeFunc & " = SUCESSO" & sNL
        
    'Exit Function
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "Exit Function" & sNL
        
    'Label
    sRotinaGrava = sRotinaGrava & sNL & "Erro_" & sNomeFunc & ":" & sNL
        
    'Erro
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & sNomeFunc & " = gerr" & sNL
        
    Call CalculaProximoErro(sProxErro)
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & "Case Else"
    sMsgErro = sMsgErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, " & """" & "ERRO_FORNECIDO_PELO_VB" & """" & ", gErr, Error, " & sProxErro & ")" & sNL
    sMsgErro = sMsgErro & sNL & TECLA_TAB & "End Select"

    'Select Case gerr
    sRotinaGrava = sRotinaGrava & sNL & sMsgErro & sNL

    'Fecha comando
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "'Fecha Comando"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "For iIndice = LBound(alComando) To UBound(alComando)"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & TECLA_TAB & "Call Comando_Fechar(alComando(iIndice))"
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "Next" & sNL

    'Exit Function
    sRotinaGrava = sRotinaGrava & sNL & TECLA_TAB & "Exit Function" & sNL
    
    'End Function
    sRotinaGrava = sRotinaGrava & sNL & "End Function"
    
    RotinaGrava_Cria = SUCESSO

    Exit Function

Erro_RotinaGrava_Cria:

    RotinaGrava_Cria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143938)

    End Select
    
    Exit Function

End Function

Private Function Telas_Cria(sDicRotinas As String, ByVal colColunasTabelas As Collection) As Long

Dim lErro As Long
Dim sNL As String
Dim objCol As ClassColunasTabelas
Dim iIndice As Integer
Dim sNomeIndice As String
Dim sChave As String
Dim iIndice2 As Integer

On Error GoTo Erro_Telas_Cria

    sNL = Chr(10)

    sDicRotinas = sDicRotinas & sNL & "INSERT INTO Telas (Nome, Projeto_Original, Classe_Original, FilialEmpresa)" & sNL & "VALUES ('" & NomeTela.Text & "', 'Telas" & ModuloTela.Text & "', 'ClassTelas" & ModuloTela.Text & "', 1)" & sNL & "GO"
    sDicRotinas = sDicRotinas & sNL & "INSERT INTO TelasModulo (SiglaModulo, NomeTela)" & sNL & "VALUES ('" & ModuloAcesso.Text & "', '" & NomeTela.Text & "')" & sNL & "GO"
    sDicRotinas = sDicRotinas & sNL & "INSERT INTO GrupoTela (CodGrupo, NomeTela, TipoDeAcesso)" & sNL & "VALUES ('supervisor', '" & NomeTela.Text & "', 1)" & sNL & "GO" & sNL

    iIndice = 0
    For Each objCol In colColunasTabelas
    
        If objCol.iChave = MARCADO Then
        
            iIndice2 = iIndice2 + 1
        
            If InStr(1, UCase(objCol.sColuna), "FILIALEMPRESA") = 0 And InStr(1, UCase(objCol.sColuna), "NUMINT") = 0 Then
                If Len(Trim(sNomeIndice)) = 0 Then sNomeIndice = objCol.sColuna
                iIndice = iIndice + 1
                sDicRotinas = sDicRotinas & sNL & "INSERT INTO TelaIndiceCampo (NomeTela, Indice, Sequencial, NomeCampo)" & sNL & "VALUES('" & NomeTela.Text & "',1," & CStr(iIndice) & ", '" & objCol.sColuna & "')" & sNL & "GO"
            End If
            
            If iIndice2 <> 1 Then
                sChave = sChave & " AND "
            End If
            
            sChave = sChave & objCol.sColuna & " = ?"
        
        End If
    
    Next
    
    If Len(Trim(sNomeIndice)) = 0 Then sNomeIndice = NomeTela.Text

    sDicRotinas = sDicRotinas & sNL & "INSERT INTO TelaIndice (NomeTela, Indice, NomeExterno)" & sNL & "VALUES('" & NomeTela.Text & "',1,'" & sNomeIndice & "')" & sNL & "GO" & sNL

    sDicRotinas = sDicRotinas & sNL & "INSERT INTO Erros (Codigo, Descricao)" & sNL & "VALUES ('AVISO_CONFIRMA_EXCLUSAO_" & UCase(NomeArq.Text) & "','Confirma a exclusão de " & NomeArq.Text & " ? ')" & sNL & "GO"

    sDicRotinas = sDicRotinas & sNL & "INSERT INTO ObjetosBD (ClasseObjeto, NomeArquivo, Tipo, SelecaoSQL, AvisaSobreposicao, NomeObjetoMsg)" & sNL & "VALUES('" & Classe.Text & "', '" & NomeArq.Text & "', 2, '" & sChave & "', 1, 'Este " & DescBrowse.Text & "')" & sNL & "GO"

    Telas_Cria = SUCESSO

    Exit Function

Erro_Telas_Cria:

    Telas_Cria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143939)

    End Select
    
    Exit Function
    
End Function

Private Function GerarTela(ByVal colColunasTabelas As Collection) As Long

Dim lErro As Long
Dim objControle As ClassCriaControles

On Error GoTo Erro_GerarTela

    Open CurDir & "\" & NomeTela.Text & ".ctm" For Output As #1

    lErro = CTX_Cria
    If lErro <> SUCESSO Then gError 131801
    
    lErro = CTL_Cria_Inicial(colColunasTabelas)
    If lErro <> SUCESSO Then gError 131802
        
    gbTemTab = False
    For Each objControle In gobjTela.colControles
        If objControle.iTipo = TIPO_FRAME Then
            gbTemTab = True
            Exit For
        End If
    Next
    
    gbTemGrid = False
    For Each objControle In gobjTela.colControles
        If objControle.iTipo = TIPO_GRID Then
            gbTemGrid = True
            Exit For
        End If
    Next
    
    lErro = CTL_Cria_Declaracao(colColunasTabelas)
    If lErro <> SUCESSO Then gError 131803

    lErro = CTL_Cria_ParteFixa
    If lErro <> SUCESSO Then gError 131804
    
    lErro = Modelo_Sub_Cria("Form_UnLoad", colColunasTabelas)
    If lErro <> SUCESSO Then gError 131808
    
    lErro = Modelo_Sub_Cria("Form_Load", colColunasTabelas)
    If lErro <> SUCESSO Then gError 131809
    
    lErro = Modelo_Function_Cria("Trata_Parametros", colColunasTabelas)
    If lErro <> SUCESSO Then gError 131810

    lErro = Modelo_Function_Cria("Move_Tela_Memoria", colColunasTabelas)
    If lErro <> SUCESSO Then gError 131811

    lErro = Modelo_Function_Cria("Tela_Extrai", colColunasTabelas)
    If lErro <> SUCESSO Then gError 131812

    lErro = Modelo_Function_Cria("Tela_Preenche", colColunasTabelas)
    If lErro <> SUCESSO Then gError 131813

    lErro = Modelo_Function_Cria("Gravar_Registro", colColunasTabelas)
    If lErro <> SUCESSO Then gError 131814
    
    lErro = Modelo_Function_Cria("Limpa_Tela_" & NomeArq.Text, colColunasTabelas)
    If lErro <> SUCESSO Then gError 131815
    
    lErro = Modelo_Function_Cria("Traz_" & NomeArq.Text & "_Tela", colColunasTabelas)
    If lErro <> SUCESSO Then gError 131816

    lErro = Modelo_Sub_Cria("BotaoGravar_Click", colColunasTabelas)
    If lErro <> SUCESSO Then gError 131817

    lErro = Modelo_Sub_Cria("BotaoFechar_Click", colColunasTabelas)
    If lErro <> SUCESSO Then gError 131818

    lErro = Modelo_Sub_Cria("BotaoLimpar_Click", colColunasTabelas)
    If lErro <> SUCESSO Then gError 131819

    lErro = Modelo_Sub_Cria("BotaoExcluir_Click", colColunasTabelas)
    If lErro <> SUCESSO Then gError 131820
    
    lErro = Cria_Scripts_DeControles(colColunasTabelas)
    If lErro <> SUCESSO Then gError 131821
    
    lErro = Cria_Scripts_DeBrowse(colColunasTabelas)
    If lErro <> SUCESSO Then gError 131822
    
    If gbTemTab Then
        lErro = Cria_Scripts_Tab()
        If lErro <> SUCESSO Then gError 131822
    End If
    
    If gbTemGrid Then
        lErro = Cria_Scripts_Grid()
        If lErro <> SUCESSO Then gError 131822
    End If

    Close #1
   
    GerarTela = SUCESSO

    Exit Function

Erro_GerarTela:

    GerarTela = gErr

    Select Case gErr
    
        Case 131801 To 131804, 131808 To 131822

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143940)

    End Select
        
    Close #1
    
    Exit Function

End Function

Private Function CTX_Cria() As Long

Dim lErro As Long
Dim sLinha As String

On Error GoTo Erro_CTX_Cria

    'Cria a referência para as figuras dos botões
    FileCopy CurDir & "\NOME_LOGICO_FISICO.ctx", CurDir & "\" & NomeTela.Text & ".ctx"
    
    CTX_Cria = SUCESSO

    Exit Function

Erro_CTX_Cria:

    CTX_Cria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143941)

    End Select
    
    Exit Function
    
End Function

Private Function CTL_Cria_Declaracao(ByVal colColunasTabelas As Collection) As Long

Dim lErro As Long
Dim objCol As ClassColunasTabelas
Dim objControle As ClassCriaControles
Dim objControleFilho As ClassCriaControles

On Error GoTo Erro_CTL_Cria_Declaracao

    Print #1, ""
    Print #1, "'Property Variables:"
    Print #1, "Dim m_Caption As String"
    Print #1, "Event Unload()"
    Print #1, ""
    Print #1, "Dim iAlterado As Integer"
    
    If gbTemTab Then
        Print #1, "Dim iFrameAtual As Integer"
    End If
    
    For Each objControle In gobjTela.colControles
        
        If objControle.iTipo = TIPO_GRID Then
            
            Print #1, ""
            Print #1, "Dim obj" & objControle.sNome & " As AdmGrid"
        
            For Each objControleFilho In objControle.colControles
                
                Print #1, "Dim iGrid_" & objControleFilho.sNome & "_Col As Integer"

            Next
        
        End If
    
    Next
    
    For Each objCol In colColunasTabelas

        If InStr(1, objCol.sColuna, "NumInt") = 0 And InStr(1, objCol.sColuna, "FilialEmpresa") = 0 And objCol.iChave = MARCADO Then

            Print #1, ""
            Print #1, "Private WithEvents objEvento" & objCol.sColuna & " As AdmEvento"
            Print #1, ""
            Exit For
                        
        End If
        
    Next
    
    CTL_Cria_Declaracao = SUCESSO

    Exit Function

Erro_CTL_Cria_Declaracao:

    CTL_Cria_Declaracao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143942)

    End Select
    
    Exit Function
    
End Function

Private Function CTL_Cria_ParteFixa() As Long

Dim lErro As Long

On Error GoTo Erro_CTL_Cria_ParteFixa

    Print #1, ""
    Print #1, "Public Function Form_Load_Ocx() As Object"
    Print #1, ""
    Print #1, "    Set Form_Load_Ocx = Me"
    Print #1, "    Caption = """ & DescBrowse.Text & """"
    Print #1, "    Call Form_Load"
    Print #1, ""
    Print #1, "End Function"
    Print #1, ""
    Print #1, "Public Function Name() As String"
    Print #1, ""
    Print #1, "    Name = """ & NomeTela.Text & """"
    Print #1, ""
    Print #1, "End Function"
    Print #1, ""
    Print #1, "Public Sub Show()"
    Print #1, "    Parent.Show"
    Print #1, "    Parent.SetFocus"
    Print #1, "End Sub"
    Print #1, ""
    Print #1, "'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!"
    Print #1, "'MappingInfo=UserControl,UserControl,-1,Controls"
    Print #1, "Public Property Get Controls() As Object"
    Print #1, "    Set Controls = UserControl.Controls"
    Print #1, "End Property"
    Print #1, ""
    Print #1, "Public Property Get hWnd() As Long"
    Print #1, "    hWnd = UserControl.hWnd"
    Print #1, "End Property"
    Print #1, ""
    Print #1, "Public Property Get Height() As Long"
    Print #1, "    Height = UserControl.Height"
    Print #1, "End Property"
    Print #1, ""
    Print #1, "Public Property Get Width() As Long"
    Print #1, "    Width = UserControl.Width"
    Print #1, "End Property"
    Print #1, ""
    Print #1, "'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!"
    Print #1, "'MappingInfo=UserControl,UserControl,-1,ActiveControl"
    Print #1, "Public Property Get ActiveControl() As Object"
    Print #1, "    Set ActiveControl = UserControl.ActiveControl"
    Print #1, "End Property"
    Print #1, ""
    Print #1, "'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!"
    Print #1, "'MappingInfo=UserControl,UserControl,-1,Enabled"
    Print #1, "Public Property Get Enabled() As Boolean"
    Print #1, "    Enabled = UserControl.Enabled"
    Print #1, "End Property"
    Print #1, ""
    Print #1, "Public Property Let Enabled(ByVal New_Enabled As Boolean)"
    Print #1, "    UserControl.Enabled() = New_Enabled"
    Print #1, "    PropertyChanged """ & "Enabled" & """"
    Print #1, "End Property"
    Print #1, ""
    Print #1, "'Load property values from storage"
    Print #1, "Private Sub UserControl_ReadProperties(PropBag As PropertyBag)"
    Print #1, "    UserControl.Enabled = PropBag.ReadProperty(" & """" & "Enabled" & """" & ", True)"
    Print #1, "End Sub"
    Print #1, ""
    Print #1, "'Write property values to storage"
    Print #1, "Private Sub UserControl_WriteProperties(PropBag As PropertyBag)"
    Print #1, "    Call PropBag.WriteProperty(" & """" & "Enabled" & """" & ", UserControl.Enabled, True)"
    Print #1, "End Sub"
    Print #1, ""
    Print #1, "Private Sub Unload(objme As Object)"
    Print #1, "   RaiseEvent Unload"
    Print #1, "End Sub"
    Print #1, ""
    Print #1, "Public Property Get Caption() As String"
    Print #1, "    Caption = m_Caption"
    Print #1, "End Property"
    Print #1, ""
    Print #1, "Public Property Let Caption(ByVal New_Caption As String)"
    Print #1, "    Parent.Caption = New_Caption"
    Print #1, "    m_Caption = New_Caption"
    Print #1, "End Property"
    Print #1, ""
    Print #1, "Public Property Get Parent() As Object"
    Print #1, "    Set Parent = UserControl.Parent"
    Print #1, "End Property"
    Print #1, "'**** fim do trecho a ser copiado *****"
    Print #1, ""
    Print #1, "Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)"
    Print #1, ""
    Print #1, "    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)"
    Print #1, ""
    Print #1, "End Sub"
    Print #1, ""
    Print #1, "Public Sub Form_Activate()"
    Print #1, ""
    Print #1, "    'Carrega os índices da tela"
    Print #1, "    Call TelaIndice_Preenche(Me)"
    Print #1, ""
    Print #1, "End Sub"
    Print #1, ""
    Print #1, "Public Sub Form_Deactivate()"
    Print #1, ""
    Print #1, "    gi_ST_SetaIgnoraClick = 1"
    Print #1, ""
    Print #1, "End Sub"
    
    CTL_Cria_ParteFixa = SUCESSO

    Exit Function

Erro_CTL_Cria_ParteFixa:

    CTL_Cria_ParteFixa = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143943)

    End Select
    
    Exit Function
    
End Function

Private Function Modelo_Function_Cria(ByVal sNomeFunc As String, ByVal colColunasTabelas As Collection) As Long

Dim lErro As Long
Dim sScriptErro As String
Dim sProxErro As String

On Error GoTo Erro_Modelo_Function_Cria

    Print #1, ""
    Print #1, "Function " & sNomeFunc & "(" & Parametros_Rotina(sNomeFunc) & ") As Long"
    Print #1, ""
    Print #1, "Dim lErro As Long"
    
    lErro = Declaracao_Rotina(sNomeFunc)
    If lErro <> SUCESSO Then gError 131804
    
    Print #1, ""
    Print #1, "On Error GoTo Erro_" & sNomeFunc
    
    lErro = Comandos_Rotina(sNomeFunc, sScriptErro, colColunasTabelas)
    If lErro <> SUCESSO Then gError 131805
    
    Print #1, ""
    Print #1, "    " & sNomeFunc & " = SUCESSO"
    Print #1, ""
    Print #1, "    Exit Function"
    Print #1, ""
    Print #1, "Erro_" & sNomeFunc & ":"
    Print #1, ""
    Print #1, "    " & sNomeFunc & " = gErr"
    
    lErro = Final1_Rotina(sNomeFunc)
    If lErro <> SUCESSO Then gError 131807
    
    Print #1, ""
    Print #1, "    Select Case gErr"
    
    lErro = Erros_Rotina(sNomeFunc, sScriptErro)
    If lErro <> SUCESSO Then gError 131806
    
    Call CalculaProximoErro(sProxErro)

    Print #1, ""
    Print #1, "        Case Else"
    Print #1, "            Call Rotina_Erro(vbOKOnly, " & """" & "ERRO_FORNECIDO_PELO_VB" & """" & ", gErr, Error, " & sProxErro & ")"
    Print #1, ""
    Print #1, "    End Select"
    
    lErro = Final2_Rotina(sNomeFunc)
    If lErro <> SUCESSO Then gError 131807
    
    Print #1, ""
    Print #1, "    Exit Function"
    Print #1, ""
    Print #1, "End Function"
    
    Modelo_Function_Cria = SUCESSO

    Exit Function

Erro_Modelo_Function_Cria:

    Modelo_Function_Cria = gErr

    Select Case gErr

        Case 131804 To 131807
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143945)

    End Select
    
    Exit Function
    
End Function

Private Function Modelo_Sub_Cria(ByVal sNomeSub As String, ByVal colColunasTabelas As Collection) As Long

Dim lErro As Long
Dim sScriptErro As String
Dim sProxErro As String

On Error GoTo Erro_Modelo_Sub_Cria

    Print #1, ""
    Print #1, "Sub " & sNomeSub & "(" & Parametros_Rotina(sNomeSub) & ")"
    Print #1, ""
    Print #1, "Dim lErro As Long"
    
    lErro = Declaracao_Rotina(sNomeSub)
    If lErro <> SUCESSO Then gError 131804
    
    Print #1, ""
    Print #1, "On Error GoTo Erro_" & sNomeSub
    
    lErro = Comandos_Rotina(sNomeSub, sScriptErro, colColunasTabelas)
    If lErro <> SUCESSO Then gError 131805
    
    Print #1, ""
    Print #1, "    Exit Sub"
    Print #1, ""
    Print #1, "Erro_" & sNomeSub & ":"
    
    lErro = Final1_Rotina(sNomeSub)
    If lErro <> SUCESSO Then gError 131807
    
    Print #1, ""
    Print #1, "    Select Case gErr"
    
    lErro = Erros_Rotina(sNomeSub, sScriptErro)
    If lErro <> SUCESSO Then gError 131806
    
    Call CalculaProximoErro(sProxErro)
    
    Print #1, ""
    Print #1, "        Case Else"
    Print #1, "            Call Rotina_Erro(vbOKOnly, " & """" & "ERRO_FORNECIDO_PELO_VB" & """" & ", gErr, Error, " & sProxErro & ")"
    Print #1, ""
    Print #1, "    End Select"
    
    lErro = Final2_Rotina(sNomeSub)
    If lErro <> SUCESSO Then gError 131807
    
    Print #1, ""
    Print #1, "    Exit Sub"
    Print #1, ""
    Print #1, "End Sub"
    
    Modelo_Sub_Cria = SUCESSO

    Exit Function

Erro_Modelo_Sub_Cria:

    Modelo_Sub_Cria = gErr

    Select Case gErr
    
        Case 131804 To 131807

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143947)

    End Select
    
    Exit Function
    
End Function

Private Function Parametros_Rotina(ByVal sNomeRot As String) As String

Dim lErro As Long
Dim sOBJ As String

On Error GoTo Erro_Parametros_Rotina

    sOBJ = "obj" & Mid(Classe.Text, 6, Len(Classe.Text) - 5)

    Select Case sNomeRot
            
        Case "Form_Load"
            Parametros_Rotina = ""
    
        Case "Form_UnLoad"
            Parametros_Rotina = "Cancel as Integer"
    
        Case "Trata_Parametros"
            Parametros_Rotina = "Optional " & sOBJ & " AS " & Classe.Text
            
        Case "Move_Tela_Memoria", "Traz_" & NomeArq.Text & "_Tela"
            Parametros_Rotina = sOBJ & " AS " & Classe.Text
        
        Case "Tela_Extrai"
            Parametros_Rotina = "sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro"
        
        Case "Tela_Preenche"
            Parametros_Rotina = "colCampoValor As AdmColCampoValor"
            
        Case "Gravar_Registro"
            Parametros_Rotina = ""
        
        Case "Limpa_Tela_" & NomeArq.Text
            Parametros_Rotina = ""
        
        Case "BotaoGravar_Click"
            Parametros_Rotina = ""
        
        Case "BotaoFechar_Click"
            Parametros_Rotina = ""
        
        Case "BotaoLimpar_Click"
            Parametros_Rotina = ""
        
        Case "BotaoExcluir_Click"
            Parametros_Rotina = ""
       
        Case Else
            Parametros_Rotina = ""
        
    End Select
    
    Exit Function

Erro_Parametros_Rotina:

    Parametros_Rotina = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143948)

    End Select
    
    Exit Function
    
End Function

Private Function Declaracao_Rotina(ByVal sNomeRot As String) As Long

Dim lErro As Long
Dim sOBJ As String

On Error GoTo Erro_Declaracao_Rotina

    sOBJ = "obj" & Mid(Classe.Text, 6, Len(Classe.Text) - 5)

    Select Case sNomeRot
    
        Case "Form_Load"
        
        Case "Trata_Parametros"
        
        Case "Move_Tela_Memoria"
        
        Case "Tela_Extrai"
            Print #1, "Dim " & sOBJ & " As New " & Classe.Text
        
        Case "Tela_Preenche"
            Print #1, "Dim " & sOBJ & " As New " & Classe.Text
        
        Case "Gravar_Registro"
            Print #1, "Dim " & sOBJ & " As New " & Classe.Text
        
        Case "Limpa_Tela_" & NomeArq.Text
        
        Case "Traz_" & NomeArq.Text & "_Tela"
        
        Case "BotaoGravar_Click"
        
        Case "BotaoFechar_Click"
        
        Case "BotaoLimpar_Click"
        
        Case "BotaoExcluir_Click"
            Print #1, "Dim " & sOBJ & " As New " & Classe.Text
            Print #1, "Dim vbMsgRes As VbMsgBoxResult"
        
        Case Else
        
    End Select
    
    Declaracao_Rotina = SUCESSO

    Exit Function

Erro_Declaracao_Rotina:

    Declaracao_Rotina = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143949)

    End Select
    
    Exit Function
    
End Function

Private Function Cria_Scripts_DeControles(ByVal colCampo As Collection) As Long

Dim lErro As Long
Dim objCol As ClassColunasTabelas

On Error GoTo Erro_Cria_Scripts_DeControles

    For Each objCol In colCampo

        If InStr(1, objCol.sColuna, "NumInt") = 0 And InStr(1, objCol.sColuna, "FilialEmpresa") = 0 Then

            Select Case objCol.sColunaTipo
            
                Case "datetime"
                    Call Cria_Scripts_DeControles_Data(objCol)
    
                Case "float"
                    Call Cria_Scripts_DeControles_Double(objCol)
            
                Case "int"
                    Call Cria_Scripts_DeControles_Long(objCol)
            
                Case "smallint"
                    Call Cria_Scripts_DeControles_Integer(objCol)
                    
                Case "varchar", "char"
                    Call Cria_Scripts_DeControles_String(objCol)
                
            End Select
            
            Call Cria_Scripts_DeControles_Alterado(objCol)
                        
        End If
        
    Next
    
    Cria_Scripts_DeControles = SUCESSO

    Exit Function

Erro_Cria_Scripts_DeControles:

    Cria_Scripts_DeControles = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143950)

    End Select
    
    Exit Function
    
End Function

Private Function Cria_Scripts_DeBrowse(ByVal colCampo As Collection) As Long

Dim lErro As Long
Dim objCol As ClassColunasTabelas
Dim sOBJBrowse As String
Dim sOBJ As String
Dim sProxErro As String
Dim sErro As String
Dim iQtdChaves As Integer
Dim sFiltro As String
Dim sNL As String

On Error GoTo Erro_Cria_Scripts_DeBrowse

    sOBJBrowse = ""
    
    sOBJ = "obj" & Mid(Classe.Text, 6, Len(Classe.Text) - 5)
    sNL = Chr(10)

    For Each objCol In colCampo
        If InStr(1, objCol.sColuna, "NumInt") = 0 And InStr(1, objCol.sColuna, "FilialEmpresa") = 0 And objCol.iChave = MARCADO Then
            sOBJBrowse = "objEvento" & objCol.sColuna
            Exit For
        End If
    Next
    
    For Each objCol In colCampo
        If objCol.iChave = MARCADO Then
            iQtdChaves = iQtdChaves + 1
            If iQtdChaves <> 1 Then sFiltro = sFiltro & ", "
            sFiltro = sFiltro & sOBJ & "." & objCol.sAtributoClasse
        End If
    Next
    
    If Len(Trim(sOBJBrowse)) > 0 Then
        
        Print #1, ""
        Print #1, "Private Sub " & sOBJBrowse & "_evSelecao(obj1 As Object)"
        Print #1, ""
        Print #1, "Dim lErro As Long"
        Print #1, "Dim " & sOBJ & " As " & Classe.Text
        Print #1, ""
        Print #1, "On Error GoTo Erro_" & sOBJBrowse & "_evSelecao"
        Print #1, ""
        Print #1, "    Set " & sOBJ & " = obj1"
        Print #1, ""
                        
        Call CalculaProximoErro(sProxErro)
        
        Print #1, "    'Mostra os dados do " & NomeArq.Text & " na tela"
        Print #1, "    lErro = Traz_" & NomeArq.Text & "_Tela(" & sOBJ & ")"
        Print #1, "    If lErro <> SUCESSO Then gError " & sProxErro
        sErro = sErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro & sNL
        
        Print #1, ""
        Print #1, "    Me.Show"
        Print #1, ""
        Print #1, "    Exit Sub"
        Print #1, ""
        Print #1, "Erro_" & sOBJBrowse & "_evSelecao:"
        Print #1, ""
        Print #1, "    Select Case gErr"
        
        Print #1, sErro
               
        Print #1, ""
        Print #1, "        Case Else"
        
        Call CalculaProximoErro(sProxErro)
        Print #1, "            Call Rotina_Erro(vbOKOnly, " & """" & "ERRO_FORNECIDO_PELO_VB" & """" & ", gErr, Error, " & sProxErro & ")"
        Print #1, ""
        Print #1, "    End Select"
        Print #1, ""
        Print #1, "    Exit Sub"
        Print #1, ""
        Print #1, "End Sub"
        
        For Each objCol In colCampo
            
            If InStr(1, objCol.sColuna, "NumInt") = 0 And InStr(1, objCol.sColuna, "FilialEmpresa") = 0 And objCol.iChave = MARCADO Then
                
                Print #1, ""
                Print #1, "Private Sub Label" & objCol.sColuna & "_Click()"
                Print #1, ""
                Print #1, "Dim lErro As Long"
                Print #1, "Dim " & sOBJ & " As New " & Classe.Text
                Print #1, "Dim colSelecao As New Collection"
                Print #1, ""
                Print #1, "On Error GoTo Erro_Label" & objCol.sColuna & "_Click"
                Print #1, ""
                Print #1, "    'Verifica se o " & objCol.sColuna & " foi preenchido"
                Print #1, "    If Len(Trim(" & objCol.sColuna & ".Text)) <> 0 Then"
                Print #1, ""
                Print #1, "        " & sOBJ & "." & objCol.sAtributoClasse & "= " & objCol.sColuna & ".Text"
                Print #1, ""
                Print #1, "    End If"
                Print #1, ""
                Print #1, "    Call Chama_Tela(" & """" & NomeBrowse.Text & """" & ", colSelecao, " & sOBJ & ", " & sOBJBrowse & ")"
                Print #1, ""
                Print #1, "    Exit Sub"
                Print #1, ""
                Print #1, "Erro_Label" & objCol.sColuna & "_Click:"
                Print #1, ""
                Print #1, "    Select Case gErr"
                Print #1, ""
                Print #1, "        Case Else"
                
                Call CalculaProximoErro(sProxErro)
                Print #1, "            Call Rotina_Erro(vbOKOnly, " & """" & "ERRO_FORNECIDO_PELO_VB" & """" & ", gErr, Error, " & sProxErro & ")"
                Print #1, ""
                Print #1, "    End Select"
                Print #1, ""
                Print #1, "    Exit Sub"
                Print #1, ""
                Print #1, "End Sub"
                
            End If
            
        Next

    End If
    
    Cria_Scripts_DeBrowse = SUCESSO

    Exit Function

Erro_Cria_Scripts_DeBrowse:

    Cria_Scripts_DeBrowse = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143953)

    End Select
    
    Exit Function
    
End Function

Private Function Cria_Scripts_DeControles_Data(ByVal objCol As ClassColunasTabelas) As Long

Dim lErro As Long
Dim sProxErro As String
Dim sErro As String

On Error GoTo Erro_Cria_Scripts_DeControles_Data
        
    Print #1, ""
    Print #1, "Private Sub UpDown" & objCol.sColuna & "_DownClick()"
    Print #1, ""
    Print #1, "Dim lErro As Long"
    Print #1, "Dim sData As String"
    
'    If objCol.iChave Then Call Cria_Scripts_DeControles_Chave1(objCol)
    
    Call CalculaProximoErro(sProxErro)
    
    Print #1, ""
    Print #1, "On Error GoTo Erro_UpDown" & objCol.sColuna & "_DownClick"
    Print #1, ""
    Print #1, "    " & objCol.sColuna & ".SetFocus"
    Print #1, ""
    Print #1, "    If Len(" & objCol.sColuna & ".ClipText) > 0 Then"
    Print #1, ""
    Print #1, "        sData = " & objCol.sColuna & ".Text"
    Print #1, ""
    Print #1, "        lErro = Data_Diminui(sData)"
    Print #1, "        If lErro <> SUCESSO Then gError " & sProxErro
    Print #1, ""
    Print #1, "        " & objCol.sColuna & ".Text = sData"
    
'    If objCol.iChave Then Call Cria_Scripts_DeControles_Chave2(sErro)
    
    Print #1, ""
    Print #1, "    End If"
    Print #1, ""
    Print #1, "    Exit Sub"
    Print #1, ""
    Print #1, "Erro_UpDown" & objCol.sColuna & "_DownClick:"
    Print #1, ""
    Print #1, "    Select Case gErr"
    
    If Len(Trim(sErro)) > 0 Then Print #1, sErro
    
    Print #1, ""
    Print #1, "        Case " & sProxErro
    Print #1, ""
    Print #1, "        Case Else"
    
    Call CalculaProximoErro(sProxErro)
    Print #1, "            Call Rotina_Erro(vbOKOnly, " & """" & "ERRO_FORNECIDO_PELO_VB" & """" & ", gErr, Error, " & sProxErro & ")"
    Print #1, ""
    Print #1, "    End Select"
    Print #1, ""
    Print #1, "    Exit Sub"
    Print #1, ""
    Print #1, "End Sub"
    
    Call CalculaProximoErro(sProxErro)
    
    Print #1, ""
    Print #1, "Private Sub UpDown" & objCol.sColuna & "_UpClick()"
    Print #1, ""
    Print #1, "Dim lErro As Long"
    Print #1, "Dim sData As String"
    Print #1, ""
    Print #1, "On Error GoTo Erro_UpDown" & objCol.sColuna & "_UpClick"
    Print #1, ""
    Print #1, "    " & objCol.sColuna & ".SetFocus"
    Print #1, ""
    Print #1, "    If Len(Trim(" & objCol.sColuna & ".ClipText)) > 0 Then"
    Print #1, ""
    Print #1, "        sData = " & objCol.sColuna & ".Text"
    Print #1, ""
    Print #1, "        lErro = Data_Aumenta(sData)"
    Print #1, "        If lErro <> SUCESSO Then gError " & sProxErro
    Print #1, ""
    Print #1, "        " & objCol.sColuna & ".Text = sData"
    Print #1, ""
    Print #1, "    End If"
    Print #1, ""
    Print #1, "    Exit Sub"
    Print #1, ""
    Print #1, "Erro_UpDown" & objCol.sColuna & "_UpClick:"
    Print #1, ""
    Print #1, "    Select Case gErr"
    Print #1, ""
    Print #1, "        Case " & sProxErro
    Print #1, ""
    Print #1, "        Case Else"
    
    Call CalculaProximoErro(sProxErro)
    Print #1, "            Call Rotina_Erro(vbOKOnly, " & """" & "ERRO_FORNECIDO_PELO_VB" & """" & ", gErr, Error, " & sProxErro & ")"
    Print #1, ""
    Print #1, "    End Select"
    Print #1, ""
    Print #1, "    Exit Sub"
    Print #1, ""
    Print #1, "End Sub"
    
    Print #1, ""
    Print #1, "Private Sub " & objCol.sColuna & "_GotFocus()"
    Print #1, "    "
    Print #1, "    Call MaskEdBox_TrataGotFocus(" & objCol.sColuna & ", iAlterado)"
    Print #1, "    "
    Print #1, "End Sub"

    Call CalculaProximoErro(sProxErro)

    Print #1, ""
    Print #1, "Private Sub " & objCol.sColuna & "_Validate(Cancel As Boolean)"
    Print #1, ""
    Print #1, "Dim lErro As Long"
    Print #1, ""
    Print #1, "On Error GoTo Erro_" & objCol.sColuna & "_Validate"
    Print #1, ""
    Print #1, "    If Len(Trim(" & objCol.sColuna & ".ClipText)) <> 0 Then "
    Print #1, ""
    Print #1, "        lErro = Data_Critica(" & objCol.sColuna & ".Text)"
    Print #1, "        If lErro <> SUCESSO Then gError " & sProxErro
    Print #1, ""
    Print #1, "    End If"
    Print #1, ""
    Print #1, "    Exit Sub"
    Print #1, ""
    Print #1, "Erro_" & objCol.sColuna & "_Validate:"
    Print #1, ""
    Print #1, "    Cancel = True"
    Print #1, ""
    Print #1, "    Select Case gErr"
    Print #1, ""
    Print #1, "        Case " & sProxErro
    Print #1, ""
    Print #1, "        Case Else"
    
    Call CalculaProximoErro(sProxErro)
    Print #1, "            Call Rotina_Erro(vbOKOnly, " & """" & "ERRO_FORNECIDO_PELO_VB" & """" & ", gErr, Error, " & sProxErro & ")"
    Print #1, ""
    Print #1, "    End Select"
    Print #1, ""
    Print #1, "    Exit Sub"
    Print #1, ""
    Print #1, "End Sub"

    Cria_Scripts_DeControles_Data = SUCESSO

    Exit Function

Erro_Cria_Scripts_DeControles_Data:

    Cria_Scripts_DeControles_Data = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143957)

    End Select
    
    Exit Function
    
End Function

Private Function Cria_Scripts_DeControles_String(ByVal objCol As ClassColunasTabelas) As Long

Dim lErro As Long
Dim sProxErro As String
Dim sErro As String

On Error GoTo Erro_Cria_Scripts_DeControles_String
    
    Print #1, ""
    Print #1, "Private Sub " & objCol.sColuna & "_Validate(Cancel As Boolean)"
    Print #1, ""
    Print #1, "Dim lErro As Long"
    
'    If objCol.iChave Then Call Cria_Scripts_DeControles_Chave1(objCol)
 
'    Call CalculaProximoErro(sProxErro)
 
    Print #1, ""
    Print #1, "On Error GoTo Erro_" & objCol.sColuna & "_Validate"
    Print #1, ""
    Print #1, "    'Verifica se " & objCol.sColuna & " está preenchida"
    Print #1, "    If Len(Trim(" & objCol.sColuna & ".Text)) <> 0 Then "
    Print #1, ""
    Print #1, "       '#######################################"
    Print #1, "       'CRITICA " & objCol.sColuna
    Print #1, "       '#######################################"
    
'    If objCol.iChave Then Call Cria_Scripts_DeControles_Chave2(sErro)
    
    Print #1, ""
    Print #1, "    End If"
    Print #1, ""
    Print #1, "    Exit Sub"
    Print #1, ""
    Print #1, "Erro_" & objCol.sColuna & "_Validate:"
    Print #1, ""
    Print #1, "    Cancel = True"
    Print #1, ""
    Print #1, "    Select Case gErr"
    Print #1, ""
    
    If Len(Trim(sErro)) > 0 Then Print #1, sErro
    
'    Print #1, "        Case " & sProxErro
'    Print #1, ""
    Print #1, "        Case Else"
    
    Call CalculaProximoErro(sProxErro)
    Print #1, "            Call Rotina_Erro(vbOKOnly, " & """" & "ERRO_FORNECIDO_PELO_VB" & """" & ", gErr, Error, " & sProxErro & ")"
    Print #1, ""
    Print #1, "    End Select"
    Print #1, ""
    Print #1, "    Exit Sub"
    Print #1, ""
    Print #1, "End Sub"
    
    Cria_Scripts_DeControles_String = SUCESSO

    Exit Function

Erro_Cria_Scripts_DeControles_String:

    Cria_Scripts_DeControles_String = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143959)

    End Select
    
    Exit Function
    
End Function

Private Function Cria_Scripts_DeControles_Double(ByVal objCol As ClassColunasTabelas) As Long

Dim lErro As Long
Dim sProxErro As String
Dim sErro As String

On Error GoTo Erro_Cria_Scripts_DeControles_Double

    Print #1, ""
    Print #1, "Private Sub " & objCol.sColuna & "_Validate(Cancel As Boolean)"
    Print #1, ""
    Print #1, "Dim lErro As Long"
    
'    If objCol.iChave Then Call Cria_Scripts_DeControles_Chave1(objCol)

    Call CalculaProximoErro(sProxErro)

    Print #1, ""
    Print #1, "On Error GoTo Erro_" & objCol.sColuna & "_Validate"
    Print #1, ""
    Print #1, "    'Verifica se " & objCol.sColuna & " está preenchida"
    Print #1, "    If Len(Trim(" & objCol.sColuna & ".Text)) <> 0 Then "
    Print #1, ""
    Print #1, "       'Critica a " & objCol.sColuna & ""
    Print #1, "       lErro = Valor_Positivo_Critica(" & objCol.sColuna & ".Text)"
    Print #1, "       If lErro <> SUCESSO Then gError " & sProxErro
    
'    If objCol.iChave Then Call Cria_Scripts_DeControles_Chave2(sErro)
    
    Print #1, ""
    Print #1, "    End If"
    Print #1, ""
    Print #1, "    Exit Sub"
    Print #1, ""
    Print #1, "Erro_" & objCol.sColuna & "_Validate:"
    Print #1, ""
    Print #1, "    Cancel = True"
    Print #1, ""
    Print #1, "    Select Case gErr"
    
    If Len(Trim(sErro)) > 0 Then Print #1, sErro
    
    Print #1, ""
    Print #1, "        Case " & sProxErro
    Print #1, ""
    Print #1, "        Case Else"
    
    Call CalculaProximoErro(sProxErro)
    Print #1, "            Call Rotina_Erro(vbOKOnly, " & """" & "ERRO_FORNECIDO_PELO_VB" & """" & ", gErr, Error, " & sProxErro & ")"
    Print #1, ""
    Print #1, "    End Select"
    Print #1, ""
    Print #1, "    Exit Sub"
    Print #1, ""
    Print #1, "End Sub"
    
    Print #1, ""
    Print #1, "Private Sub " & objCol.sColuna & "_GotFocus()"
    Print #1, "    "
    Print #1, "    Call MaskEdBox_TrataGotFocus(" & objCol.sColuna & ", iAlterado)"
    Print #1, "    "
    Print #1, "End Sub"
    
    Cria_Scripts_DeControles_Double = SUCESSO

    Exit Function

Erro_Cria_Scripts_DeControles_Double:

    Cria_Scripts_DeControles_Double = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143961)

    End Select
    
    Exit Function
    
End Function

Private Function Cria_Scripts_DeControles_Alterado(ByVal objCol As ClassColunasTabelas) As Long

Dim lErro As Long

On Error GoTo Erro_Cria_Scripts_DeControles_Alterado
      
    Print #1, ""
    Print #1, "Private Sub " & objCol.sColuna & "_Change()"
    Print #1, "    iAlterado = REGISTRO_ALTERADO"
    Print #1, "End Sub"

    Cria_Scripts_DeControles_Alterado = SUCESSO

    Exit Function

Erro_Cria_Scripts_DeControles_Alterado:

    Cria_Scripts_DeControles_Alterado = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143962)

    End Select
    
    Exit Function
    
End Function

Private Function Cria_Scripts_DeControles_Chave1(ByVal objCol As ClassColunasTabelas) As Long

Dim lErro As Long
Dim sOBJ As String

On Error GoTo Erro_Cria_Scripts_DeControles_Chave1
      
    sOBJ = "obj" & Mid(Classe.Text, 6, Len(Classe.Text) - 5)

    Print #1, "Dim " & sOBJ & " As New " & Classe.Text

    Cria_Scripts_DeControles_Chave1 = SUCESSO

    Exit Function

Erro_Cria_Scripts_DeControles_Chave1:

    Cria_Scripts_DeControles_Chave1 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143963)

    End Select
    
    Exit Function
    
End Function

Private Function Cria_Scripts_DeControles_Chave2(sErro As String) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objColunasTabelas As ClassColunasTabelas
Dim colColunasTabelas As New Collection
Dim sSigla As String
Dim iTipo As Integer
Dim sTipoVB As String
Dim sOBJ As String
Dim sNL As String
Dim sProxErro As String

On Error GoTo Erro_Cria_Scripts_DeControles_Chave2

    sNL = Chr(10)
    sOBJ = "obj" & Mid(Classe.Text, 6, Len(Classe.Text) - 5)

    lErro = Move_Tela_Memoria(colColunasTabelas)
    If lErro <> SUCESSO Then gError 131822

    Print #1, ""

    iIndice = 0
    For Each objColunasTabelas In colColunasTabelas

        iIndice = iIndice + 1

        With objColunasTabelas

            If InStr(1, .sColuna, "NumInt") = 0 And InStr(1, .sColuna, "FilialEmpresa") = 0 And .iChave = MARCADO Then

                Call ObtemSiglaTipo(.sColunaTipo, sSigla, iTipo, sTipoVB)

                Select Case sSigla

                    Case "s"
                        Print #1, TECLA_TAB & TECLA_TAB & sOBJ & "." & .sAtributoClasse & " = " & .sColuna & ".text"

                    Case "i"
                        Print #1, TECLA_TAB & TECLA_TAB & sOBJ & "." & .sAtributoClasse & " = StrParaInt(" & .sColuna & ".text)"

                    Case "dt"
                        Print #1, TECLA_TAB & TECLA_TAB & sOBJ & "." & .sAtributoClasse & " = StrParaDate(" & .sColuna & ".text)"

                    Case "d"
                        Print #1, TECLA_TAB & TECLA_TAB & sOBJ & "." & .sAtributoClasse & " = StrParaDbl(" & .sColuna & ".text)"

                    Case "l"
                        Print #1, TECLA_TAB & TECLA_TAB & sOBJ & "." & .sAtributoClasse & " = StrParaLong(" & .sColuna & ".text)"

                End Select

            End If

            If InStr(1, .sColuna, "FilialEmpresa") <> 0 And .iChave = MARCADO Then
                Print #1, TECLA_TAB & TECLA_TAB & sOBJ & "." & .sAtributoClasse & " = giFilialEmpresa"
            End If

        End With

    Next

    Call CalculaProximoErro(sProxErro)

    Print #1, ""
    Print #1, TECLA_TAB & TECLA_TAB & "lErro = Traz_" & NomeArq.Text & "_Tela(" & sOBJ & ")"
    Print #1, TECLA_TAB & TECLA_TAB & "If lErro <> SUCESSO then gError " & sProxErro

    sErro = sErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro

    Cria_Scripts_DeControles_Chave2 = SUCESSO

    Exit Function

Erro_Cria_Scripts_DeControles_Chave2:

    Cria_Scripts_DeControles_Chave2 = gErr

    Select Case gErr
    
        Case 131822

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143964)

    End Select
    
    Exit Function
    
End Function

Private Function Cria_Scripts_DeControles_Long(ByVal objCol As ClassColunasTabelas) As Long

Dim lErro As Long
Dim sProxErro As String
Dim sErro As String

On Error GoTo Erro_Cria_Scripts_DeControles_Long
               
    Print #1, ""
    Print #1, "Private Sub " & objCol.sColuna & "_Validate(Cancel As Boolean)"
    Print #1, ""
    Print #1, "Dim lErro As Long"
    
'    If objCol.iChave Then Call Cria_Scripts_DeControles_Chave1(objCol)
    
    Call CalculaProximoErro(sProxErro)
    
    Print #1, ""
    Print #1, "On Error GoTo Erro_" & objCol.sColuna & "_Validate"
    Print #1, ""
    Print #1, "    'Verifica se " & objCol.sColuna & " está preenchida"
    Print #1, "    If Len(Trim(" & objCol.sColuna & ".Text)) <> 0 Then "
    Print #1, ""
    Print #1, "       'Critica a " & objCol.sColuna & ""
    Print #1, "       lErro = Long_Critica(" & objCol.sColuna & ".Text)"
    Print #1, "       If lErro <> SUCESSO Then gError " & sProxErro
    
'    If objCol.iChave Then Call Cria_Scripts_DeControles_Chave2(sErro)
    
    Print #1, ""
    Print #1, "    End If"
    Print #1, ""
    Print #1, "    Exit Sub"
    Print #1, ""
    Print #1, "Erro_" & objCol.sColuna & "_Validate:"
    Print #1, ""
    Print #1, "    Cancel = True"
    Print #1, ""
    Print #1, "    Select Case gErr"
    
    If Len(Trim(sErro)) > 0 Then Print #1, sErro
    
    Print #1, ""
    Print #1, "        Case " & sProxErro
    Print #1, ""
    Print #1, "        Case Else"
    
    Call CalculaProximoErro(sProxErro)
    Print #1, "            Call Rotina_Erro(vbOKOnly, " & """" & "ERRO_FORNECIDO_PELO_VB" & """" & ", gErr, Error, " & sProxErro & ")"
    Print #1, ""
    Print #1, "    End Select"
    Print #1, ""
    Print #1, "    Exit Sub"
    Print #1, ""
    Print #1, "End Sub"
    
    Print #1, ""
    Print #1, "Private Sub " & objCol.sColuna & "_GotFocus()"
    Print #1, "    "
    Print #1, "    Call MaskEdBox_TrataGotFocus(" & objCol.sColuna & ", iAlterado)"
    Print #1, "    "
    Print #1, "End Sub"
    
    Cria_Scripts_DeControles_Long = SUCESSO

    Exit Function

Erro_Cria_Scripts_DeControles_Long:

    Cria_Scripts_DeControles_Long = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143966)

    End Select
    
    Exit Function
    
End Function

Private Function Cria_Scripts_DeControles_Integer(ByVal objCol As ClassColunasTabelas) As Long

Dim lErro As Long
Dim sProxErro As String
Dim sErro As String

On Error GoTo Erro_Cria_Scripts_DeControles_Integer

    Print #1, ""
    Print #1, "Private Sub " & objCol.sColuna & "_Validate(Cancel As Boolean)"
    Print #1, ""
    Print #1, "Dim lErro As Long"
    
'    If objCol.iChave Then Call Cria_Scripts_DeControles_Chave1(objCol)
    
    Call CalculaProximoErro(sProxErro)
    
    Print #1, ""
    Print #1, "On Error GoTo Erro_" & objCol.sColuna & "_Validate"
    Print #1, ""
    Print #1, "    'Verifica se " & objCol.sColuna & " está preenchida"
    Print #1, "    If Len(Trim(" & objCol.sColuna & ".Text)) <> 0 Then "
    Print #1, ""
    Print #1, "       'Critica a " & objCol.sColuna & ""
    Print #1, "       lErro = Inteiro_Critica(" & objCol.sColuna & ".Text)"
    Print #1, "       If lErro <> SUCESSO Then gError " & sProxErro
    
'    If objCol.iChave Then Call Cria_Scripts_DeControles_Chave2(sErro)
    
    Print #1, ""
    Print #1, "    End If"
    Print #1, ""
    Print #1, "    Exit Sub"
    Print #1, ""
    Print #1, "Erro_" & objCol.sColuna & "_Validate:"
    Print #1, ""
    Print #1, "    Cancel = True"
    Print #1, ""
    Print #1, "    Select Case gErr"
    
    If Len(Trim(sErro)) > 0 Then Print #1, sErro
    
    Print #1, ""
    Print #1, "        Case " & sProxErro
    Print #1, ""
    Print #1, "        Case Else"
    
    Call CalculaProximoErro(sProxErro)
    Print #1, "            Call Rotina_Erro(vbOKOnly, " & """" & "ERRO_FORNECIDO_PELO_VB" & """" & ", gErr, Error, " & sProxErro & ")"
    Print #1, ""
    Print #1, "    End Select"
    Print #1, ""
    Print #1, "    Exit Sub"
    Print #1, ""
    Print #1, "End Sub"
    
    Print #1, ""
    Print #1, "Private Sub " & objCol.sColuna & "_GotFocus()"
    Print #1, "    "
    Print #1, "    Call MaskEdBox_TrataGotFocus(" & objCol.sColuna & ", iAlterado)"
    Print #1, "    "
    Print #1, "End Sub"
    
    Cria_Scripts_DeControles_Integer = SUCESSO

    Exit Function

Erro_Cria_Scripts_DeControles_Integer:

    Cria_Scripts_DeControles_Integer = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143968)

    End Select
    
    Exit Function
    
End Function

Private Function Controles_Tela_Cria(ByVal colCampo As Collection) As Long

Dim objColunasTabelas As ClassColunasTabelas
Dim lErro As Long
Dim lTop As Long
Dim iIndex As Integer
Dim lWidth As Long
Dim objControle As ClassCriaControles
Dim objControleAux As ClassCriaControles
Dim sNL As String
Dim bAchou As Boolean
Dim objTela As New ClassCriaTela
Dim bTemTab As Boolean
Dim iContTabs As Integer
Dim iIndiceAux As Integer
Dim lTopUltimo As Long

On Error GoTo Erro_Controles_Tela_Cria

    lTopUltimo = 300
    
    For Each objControle In gobjTela.colControles
        If objControle.iTipo = TIPO_FRAME Then
            objControle.lTopUltimo = 300
        End If
    Next
    
    iIndex = 5
    sNL = Chr(10) & Chr(13) & Chr(10) & Chr(13)
    
    Set objTela = New ClassCriaTela

    For Each objColunasTabelas In colCampo
    
        'Só cria o controle se ele não for o NumInt ou o FilialEmpresa
        If InStr(1, objColunasTabelas.sColuna, "NumInt") = 0 And InStr(1, objColunasTabelas.sColuna, "FilialEmpresa") = 0 Then
            
            'Para campos com tamanho maior que 50 usa um textbox, senão um MaskEditBox
            If objColunasTabelas.lColunaTamanho <= 50 Then
            
                'Calcula o Tamanho do Campo
                If objColunasTabelas.sColunaTipo = "varchar" Or objColunasTabelas.sColunaTipo = "char" Then
                    lWidth = 110 * objColunasTabelas.lColunaTamanho
                End If
                If objColunasTabelas.sColunaTipo = "int" Or objColunasTabelas.sColunaTipo = "float" Then
                    lWidth = 110 * 8
                End If
                If objColunasTabelas.sColunaTipo = "smallint" Then
                    lWidth = 110 * 5
                End If
                
                'Se o campo for do tipo data => Cria UpDown e MaskEdit Apropriados
                If objColunasTabelas.sColunaTipo = "datetime" Then
                
                    iIndex = iIndex + 1
                
                    bAchou = False
                    For Each objControle In gobjTela.colControles
                        If objControle.sNome = objColunasTabelas.sColuna Then
                            bAchou = True
                            Exit For
                        End If
                    Next
                    If Not bAchou Then
                        Set objControle = New ClassCriaControles
                        objControle.iOrdem = iIndex
                        objControle.sNome = objColunasTabelas.sColuna
                        objControle.sTipo = "MaskEdBox"
                        objControle.iTipo = TIPO_OUTRO
                    End If
                    
                    'obtém a altura do controle
                    If objControle.sFrame <> "" Then
                        For Each objControleAux In gobjTela.colControles
                            If objControleAux.sNome = objControle.sFrame And objControleAux.iTipo = TIPO_FRAME Then
                                lTop = objControleAux.lTopUltimo
                                objControleAux.lTopUltimo = objControleAux.lTopUltimo + 450
                            End If
                        Next
                    Else
                        lTop = lTopUltimo
                        lTopUltimo = lTopUltimo + 450
                    End If
                    
                    Call Limpa_ObjControle(objControle)
                
                    objControle.sScript(1) = "   Begin MSMask.MaskEdBox " & objColunasTabelas.sColuna
                    objControle.sScript(2) = "      Height          =   315"
                    objControle.sScript(3) = "      Left            =   2000"
                    objControle.sScript(4) = "      TabIndex        =   " & CStr(iIndex)
                    objControle.sScript(5) = "      Top             =   " & CStr(lTop)
                    objControle.sScript(6) = "      Width           =   1300"
                    objControle.sScript(7) = "      _ExtentX        =   2355"
                    objControle.sScript(8) = "      _ExtentY        =   556"
                    objControle.sScript(9) = "      _Version        =   393216"
                    objControle.sScript(10) = "      MaxLength      =   8"
                    objControle.sScript(11) = "      Format         =   " & """" & "dd/mm/yyyy" & """"
                    objControle.sScript(12) = "      Mask           =   " & """" & "##/##/##" & """"
                    objControle.sScript(13) = "      PromptChar     =   " & """" & " " & """"
                    objControle.sScript(14) = "   End"
                    
                    objTela.colControles.Add objControle
                    
                    iIndex = iIndex + 1
                    
                    bAchou = False
                    For Each objControle In gobjTela.colControles
                        If objControle.sNome = "UpDown" & objColunasTabelas.sColuna Then
                            bAchou = True
                            Exit For
                        End If
                    Next
                    If Not bAchou Then
                        Set objControle = New ClassCriaControles
                        objControle.iOrdem = iIndex
                        objControle.sNome = "UpDown" & objColunasTabelas.sColuna
                        objControle.sTipo = "UpDown"
                        objControle.iTipo = TIPO_OUTRO
                    End If
                    
                    Call Limpa_ObjControle(objControle)
                    
                    objControle.sScript(1) = "   Begin MSComCtl2.UpDown UpDown" & objColunasTabelas.sColuna
                    objControle.sScript(2) = "      Height          =   300"
                    objControle.sScript(3) = "      Left            =   3310"
                    objControle.sScript(4) = "      TabIndex        =   " & CStr(iIndex)
                    objControle.sScript(5) = "      TabStop         =   0             'False"
                    objControle.sScript(6) = "      Top             =   " & CStr(lTop)
                    objControle.sScript(7) = "      Width           =   240"
                    objControle.sScript(8) = "      _ExtentX        =   423"
                    objControle.sScript(9) = "      _ExtentY        =   529"
                    objControle.sScript(10) = "      _Version        =   393216"
                    objControle.sScript(11) = "      Enabled         =   -1            'True"
                    objControle.sScript(12) = "   End"
                
                    objTela.colControles.Add objControle
               
                Else 'MaskEditBox
                
                    iIndex = iIndex + 1
                    
                    bAchou = False
                    For Each objControle In gobjTela.colControles
                        If objControle.sNome = objColunasTabelas.sColuna Then
                            bAchou = True
                            Exit For
                        End If
                    Next
                    If Not bAchou Then
                        Set objControle = New ClassCriaControles
                        objControle.iOrdem = iIndex
                        objControle.sNome = objColunasTabelas.sColuna
                        objControle.sTipo = "MaskEdBox"
                        objControle.iTipo = TIPO_OUTRO
                    End If
                    
                    'obtém a altura do controle
                    If objControle.sFrame <> "" Then
                        For Each objControleAux In gobjTela.colControles
                            If objControleAux.sNome = objControle.sFrame And objControleAux.iTipo = TIPO_FRAME Then
                                lTop = objControleAux.lTopUltimo
                                objControleAux.lTopUltimo = objControleAux.lTopUltimo + 450
                            End If
                        Next
                    Else
                        lTop = lTopUltimo
                        lTopUltimo = lTopUltimo + 450
                    End If
                    
                    Call Limpa_ObjControle(objControle)
                    
                    objControle.sScript(1) = "   Begin MSMask.MaskEdBox " & objColunasTabelas.sColuna
                    objControle.sScript(2) = "      Height          =   315"
                    objControle.sScript(3) = "      Left            =   2000"
                    objControle.sScript(4) = "      TabIndex        =   " & iIndex
                    objControle.sScript(5) = "      Top             =   " & CStr(lTop)
                    objControle.sScript(6) = "      Width           =   " & CStr(lWidth)
                    objControle.sScript(7) = "      _ExtentX        =   2699"
                    objControle.sScript(8) = "      _ExtentY        =   661"
                    objControle.sScript(9) = "      _Version        =   393216"
                    objControle.sScript(10) = "      MaxLength       =   " & objColunasTabelas.lColunaTamanho
                    objControle.sScript(11) = "      PromptChar      =   " & """" & " " & """"
                    objControle.sScript(12) = "   End"
            
                    objTela.colControles.Add objControle
           
                End If
            
            Else 'TextBox
                
                iIndex = iIndex + 1
                
                bAchou = False
                For Each objControle In gobjTela.colControles
                    If objControle.sNome = objColunasTabelas.sColuna Then
                        bAchou = True
                        Exit For
                    End If
                Next
                If Not bAchou Then
                    Set objControle = New ClassCriaControles
                    objControle.iOrdem = iIndex
                    objControle.sNome = objColunasTabelas.sColuna
                    objControle.sTipo = "TextBox"
                    objControle.iTipo = TIPO_OUTRO
                End If
                
                'obtém a altura do controle
                If objControle.sFrame <> "" Then
                    For Each objControleAux In gobjTela.colControles
                        If objControleAux.sNome = objControle.sFrame And objControleAux.iTipo = TIPO_FRAME Then
                            lTop = objControleAux.lTopUltimo
                            objControleAux.lTopUltimo = objControleAux.lTopUltimo + 450
                        End If
                    Next
                Else
                    lTop = lTopUltimo
                    lTopUltimo = lTopUltimo + 450
                End If
                    
                Call Limpa_ObjControle(objControle)
            
                objControle.sScript(1) = "   Begin VB.TextBox " & objColunasTabelas.sColuna
                objControle.sScript(2) = "      Height          =   315"
                objControle.sScript(3) = "      Left            =   2000"
                objControle.sScript(4) = "      MaxLength       =   " & objColunasTabelas.lColunaTamanho
                objControle.sScript(5) = "      TabIndex        =   " & iIndex
                objControle.sScript(6) = "      Text            =   " & """" & """" & ""
                objControle.sScript(7) = "      Top             =   " & CStr(lTop)
                objControle.sScript(8) = "      Width           =   5500"
                objControle.sScript(9) = "   End"
            
                objTela.colControles.Add objControle
            
            End If
            
            iIndex = iIndex + 1
            
            bAchou = False
            For Each objControle In gobjTela.colControles
                If objControle.sNome = "Label" & objColunasTabelas.sColuna Then
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then
                Set objControle = New ClassCriaControles
                objControle.iOrdem = 0
                objControle.sNome = "Label" & objColunasTabelas.sColuna
                objControle.sTipo = "Label"
                objControle.iTipo = TIPO_OUTRO
            End If
                
            Call Limpa_ObjControle(objControle)
                
            'Label
            objControle.sScript(1) = "   Begin VB.Label Label" & objColunasTabelas.sColuna
            objControle.sScript(2) = "      Alignment       =   1  'Right Justify"
            objControle.sScript(3) = "      Caption         =   " & """" & objColunasTabelas.sDescricao & ":" & """"
            objControle.sScript(4) = "      BeginProperty Font"
            objControle.sScript(5) = "         Name            = " & """" & "MS Sans Serif" & """"
            objControle.sScript(6) = "         Size            =   8.25"
            objControle.sScript(7) = "         Charset         =   0"
            objControle.sScript(8) = "         Weight          =   700"
            objControle.sScript(9) = "         Underline       =   0              'False"
            objControle.sScript(10) = "         Italic          =   0              'False"
            objControle.sScript(11) = "         Strikethrough   =   0              'False"
            objControle.sScript(12) = "      EndProperty"
            
            iIndiceAux = 12
            If objColunasTabelas.iChave = MARCADO Then
                iIndiceAux = iIndiceAux + 1
                objControle.sScript(iIndiceAux) = "      ForeColor       =   &H00000080&"
            End If

            objControle.sScript(iIndiceAux + 1) = "      Height          =   315"
            objControle.sScript(iIndiceAux + 2) = "      Left            =   375"
            
            If objColunasTabelas.iChave = MARCADO Then
                iIndiceAux = iIndiceAux + 1
                objControle.sScript(iIndiceAux + 2) = "      MousePointer    = 14       'Arrow and Question"
            End If

            objControle.sScript(iIndiceAux + 3) = "      TabIndex        = " & iIndex
            objControle.sScript(iIndiceAux + 4) = "      Top             = " & CStr(lTop + 25)
            objControle.sScript(iIndiceAux + 5) = "      Width           = 1500"
            objControle.sScript(iIndiceAux + 6) = "   End"
            
            objTela.colControles.Add objControle
    
        End If
            
    Next
    
    bTemTab = False
    For Each objControle In gobjTela.colControles
    
        iIndex = iIndex + 1
    
        Select Case objControle.iTipo
        
            Case TIPO_FRAME
                bTemTab = True
                iContTabs = iContTabs + 1
          
                Call Limpa_ObjControle(objControle)
          
                objControle.sScript(1) = "   Begin VB.Frame FrameOpcao"
                objControle.sScript(2) = "      BorderStyle     =   0        'None"
                objControle.sScript(3) = "      Height          =   5220"
                objControle.sScript(4) = "      Index           =   " & CStr(objControle.iOrdem)
                objControle.sScript(5) = "      Left            =   135"
                objControle.sScript(6) = "      TabIndex        =   " & CStr(iIndex)
                objControle.sScript(7) = "      Top             =   660"
                
                iIndiceAux = 8
                If objControle.iOrdem <> 1 Then
                    objControle.sScript(iIndiceAux) = "      Visible         = 0          'False"
                    iIndiceAux = iIndiceAux + 1
                End If
                
                objControle.sScript(iIndiceAux) = "      Width           =   9195"
                        
                objTela.colControles.Add objControle
                       
            Case TIPO_GRID
            
                iIndex = iIndex + 1
        
                Call Limpa_ObjControle(objControle)
        
                objControle.sScript(1) = "   Begin VB.Frame Frame" & objControle.sNome
                objControle.sScript(2) = "      Caption         =   " & """" & objControle.sNome & """"
                objControle.sScript(3) = "      Height          =   2745"
                objControle.sScript(4) = "      Left            =   60"
                objControle.sScript(5) = "      TabIndex        =   " & CStr(iIndex - 1)
                objControle.sScript(6) = "      Top             =   2220"
                objControle.sScript(7) = "      Width           =   9090"
        
                iIndiceAux = 8
        
                'Controles associados as grid
                For Each objControleAux In objControle.colControles
                
                    iIndex = iIndex + 1

                    objControle.sScript(iIndiceAux) = "      Begin MSMask.MaskEdBox " & objControleAux.sNome
                    objControle.sScript(iIndiceAux + 1) = "         Height          =   315"
                    objControle.sScript(iIndiceAux + 2) = "         Left            =   315"
                    objControle.sScript(iIndiceAux + 3) = "         TabIndex        =   " & iIndex
                    objControle.sScript(iIndiceAux + 4) = "         Top             =   285"
                    objControle.sScript(iIndiceAux + 5) = "         Width           =   1500"
                    objControle.sScript(iIndiceAux + 6) = "         _ExtentX        =   2699"
                    objControle.sScript(iIndiceAux + 7) = "         _ExtentY        =   661"
                    objControle.sScript(iIndiceAux + 8) = "         _Version        =   393216"
                    objControle.sScript(iIndiceAux + 9) = "         PromptChar      =   " & """" & " " & """"
                    objControle.sScript(iIndiceAux + 10) = "      End"
                
                    iIndiceAux = iIndiceAux + 11
                
                Next
        
                objControle.sScript(iIndiceAux) = "      Begin MSFlexGridLib.MSFlexGrid " & objControle.sNome
                objControle.sScript(iIndiceAux + 1) = "         Height          =   2325"
                objControle.sScript(iIndiceAux + 2) = "         Left            =   120"
                objControle.sScript(iIndiceAux + 3) = "         TabIndex        =   " & CStr(iIndex)
                objControle.sScript(iIndiceAux + 4) = "         Top             =   285"
                objControle.sScript(iIndiceAux + 5) = "         Width           =   8820"
                objControle.sScript(iIndiceAux + 6) = "         _ExtentX        =   15558"
                objControle.sScript(iIndiceAux + 7) = "         _ExtentY        =   4075"
                objControle.sScript(iIndiceAux + 8) = "         _Version        =   393216"
                objControle.sScript(iIndiceAux + 9) = "         Rows            =   21"
                objControle.sScript(iIndiceAux + 10) = "         Cols            =   4"
                objControle.sScript(iIndiceAux + 11) = "         BackColorSel    =   -2147483643"
                objControle.sScript(iIndiceAux + 12) = "         ForeColorSel    =   -2147483640"
                objControle.sScript(iIndiceAux + 13) = "         AllowBigSelection = 0    'False"
                objControle.sScript(iIndiceAux + 14) = "         FocusRect       =   2"
                objControle.sScript(iIndiceAux + 15) = "   End"
                objControle.sScript(iIndiceAux + 16) = "End"
        
                objTela.colControles.Add objControle
        
        End Select
            
    Next
       
    'Insere os Frames e os controles dentro dos frame
    For Each objControle In objTela.colControles
    
        If objControle.iTipo = TIPO_FRAME Then
        
            Call Imprime_ObjControle(objControle)

            For Each objControleAux In objTela.colControles
        
                If objControleAux.sFrame = objControle.sNome Then
        
                    Call Imprime_ObjControle(objControleAux, "   ")
        
                End If
        
            Next
            
            Print #1, "   End"
            
        End If

    Next
    
    'Insere os controles que estão fora dos frames
    For Each objControle In objTela.colControles
    
        If objControle.iTipo <> TIPO_FRAME And objControle.sFrame = "" Then
        
            Call Imprime_ObjControle(objControle)
            
        End If

    Next
    
    'Insere o Tab
    If bTemTab Then
    
        Print #1, "   Begin MSComctlLib.TabStrip Opcao"
        Print #1, "      Height          =   5670"
        Print #1, "      Left            =   75"
        Print #1, "      TabIndex        =   0"
        Print #1, "      Top             =   255"
        Print #1, "      Width           =   9345"
        Print #1, "      _ExtentX        =   16484"
        Print #1, "      _ExtentY        =   10001"
        Print #1, "      _Version        =   393216"
        Print #1, "      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628}"
        Print #1, "         NumTabs         =   " & CStr(iContTabs)
        
        For Each objControle In objTela.colControles
        
            If objControle.iTipo = TIPO_FRAME Then
        
                Print #1, "         BeginProperty Tab" & objControle.iOrdem & " {1EFB659A-857C-11D1-B16A-00C0F0283628}"
                Print #1, "            Caption         =   " & """" & objControle.sNome & """"
                Print #1, "            ImageVarType    =   2"
                Print #1, "         EndProperty"
        
            End If
        
        Next
        
        Print #1, "      EndProperty"
        Print #1, "      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}"
        Print #1, "         Name            =   " & """" & " MS Sans Serif" & """"
        Print #1, "         Size            =   8.25"
        Print #1, "         Charset         =   0"
        Print #1, "         Weight          =   700"
        Print #1, "         Underline       =   0           'False"
        Print #1, "         Italic          =   0              'False"
        Print #1, "         Strikethrough   =   0       'False"
        Print #1, "      EndProperty"
        Print #1, "   End"
    
    End If
    
    Set gobjTela = objTela

    Controles_Tela_Cria = SUCESSO

    Exit Function

Erro_Controles_Tela_Cria:

    Controles_Tela_Cria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143969)

    End Select
    
    Exit Function
    
End Function

Private Function Comandos_Rotina(ByVal sNomeRot As String, sScriptErro As String, ByVal colColunasTabelas As Collection) As Long

Dim lErro As Long
Dim sOBJ As String
Dim sProxErro As String
Dim sNL As String
Dim objColunasTabelas As ClassColunasTabelas
Dim iIndice  As Integer
Dim sSigla As String
Dim iTipo As Integer
Dim sTipoVB As String
Dim sErroLeitura As String
Dim sAux As String
Dim sChave As String
Dim objControle As ClassCriaControles

On Error GoTo Erro_Comandos_Rotina

    sNL = Chr(10)
    sOBJ = "obj" & Mid(Classe.Text, 6, Len(Classe.Text) - 5)

    If Len(Trim(gsErroLeitura)) > 0 Then
        sErroLeitura = gsErroLeitura
    Else
        sErroLeitura = "ERRO_LEITURA_SEM_DADOS"
    End If

    Select Case sNomeRot
    
        Case "Form_UnLoad"
        
            Print #1, ""
            
            If gbTemTab Then
                Print #1, TECLA_TAB & "iFrameAtual = 1"
                Print #1, ""
            End If
            
            For Each objControle In gobjTela.colControles
            
                If objControle.iTipo = TIPO_GRID Then

                    Print #1, "    Set obj" & objControle.sNome & " = Nothing"
                
                End If
            
            Next
            
            If gbTemGrid Then Print #1, ""
            
            For Each objColunasTabelas In colColunasTabelas
        
                If InStr(1, objColunasTabelas.sColuna, "NumInt") = 0 And InStr(1, objColunasTabelas.sColuna, "FilialEmpresa") = 0 Then
        
                    Print #1, TECLA_TAB & "Set objEvento" & objColunasTabelas.sColuna & " = Nothing"
                    Exit For
                                
                End If
                
            Next
            
            Print #1, TECLA_TAB & "Call ComandoSeta_Liberar(Me.Name)"
   
        Case "Form_Load"

            Print #1, ""
            
            For Each objColunasTabelas In colColunasTabelas
        
                If InStr(1, objColunasTabelas.sColuna, "NumInt") = 0 And InStr(1, objColunasTabelas.sColuna, "FilialEmpresa") = 0 Then
        
                    Print #1, TECLA_TAB & "Set objEvento" & objColunasTabelas.sColuna & " = New AdmEvento"
                    Exit For
                                
                End If
                
            Next
            
            If gbTemGrid Then
                sScriptErro = sNL & TECLA_TAB & TECLA_TAB & "Case "
            End If
            
            For Each objControle In gobjTela.colControles
            
                If objControle.iTipo = TIPO_GRID Then
                
                    Call CalculaProximoErro(sProxErro)
                
                    Print #1, ""
                    Print #1, "    lErro = Inicializa_" & objControle.sNome & "(obj" & objControle.sNome & ")"
                    Print #1, "    If lErro <> SUCESSO Then gError " & sProxErro
                
                    sScriptErro = sScriptErro & sProxErro & ", "
                
                End If
            
            Next
            
            'Tira o último ", "
            If Len(sScriptErro) > 2 Then
                sScriptErro = left(sScriptErro, Len(sScriptErro) - 2)
            End If
            
            If gbTemTab Then
                Print #1, ""
                Print #1, "    iFrameAtual = 1"
            End If
            
            Print #1, ""
            Print #1, "    iAlterado = 0"
            Print #1, ""
            Print #1, "    lErro_Chama_Tela = SUCESSO"
    
        Case "Trata_Parametros"
            
            Call CalculaProximoErro(sProxErro)

            Print #1, ""
            Print #1, "    If Not (" & sOBJ & " Is Nothing) Then"
            Print #1, ""
            Print #1, "        lErro = Traz_" & NomeArq.Text & "_Tela(" & sOBJ & ")"
            Print #1, "        If lErro <> SUCESSO Then gError " & sProxErro
            Print #1, ""
            Print #1, "    End If"
            Print #1, ""
            Print #1, "    iAlterado = 0"
            
            sScriptErro = sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro
        
        Case "Move_Tela_Memoria"
            
            Print #1, ""
            
            iIndice = 0
            For Each objColunasTabelas In colColunasTabelas
            
                iIndice = iIndice + 1
                
                With objColunasTabelas
                
                    If InStr(1, .sColuna, "NumInt") = 0 And InStr(1, .sColuna, "FilialEmpresa") = 0 Then
        
                        Call ObtemSiglaTipo(.sColunaTipo, sSigla, iTipo, sTipoVB)
            
                        Select Case sSigla
                        
                            Case "s"
                                Print #1, TECLA_TAB & sOBJ & "." & .sAtributoClasse & " = " & .sColuna & ".text"
                            
                            Case "i"
                                Print #1, TECLA_TAB & sOBJ & "." & .sAtributoClasse & " = StrParaInt(" & .sColuna & ".text)"
                            
                            Case "dt"
                                Print #1, TECLA_TAB & "if len(trim(" & .sColuna & ".ClipText))<>0 then " & sOBJ & "." & .sAtributoClasse & " = strparadate(" & .sColuna & ".text)"
                            
                            Case "d"
                                Print #1, TECLA_TAB & sOBJ & "." & .sAtributoClasse & " = StrParaDbl(" & .sColuna & ".text)"
                    
                            Case "l"
                                Print #1, TECLA_TAB & sOBJ & "." & .sAtributoClasse & " = StrParaLong(" & .sColuna & ".text)"
                    
                        End Select
                        
                    End If
                    
                    If InStr(1, .sColuna, "FilialEmpresa") <> 0 Then
                        Print #1, TECLA_TAB & sOBJ & "." & .sAtributoClasse & " = giFilialEmpresa"
                    End If
                
                End With
            
            Next
        
        Case "Tela_Extrai"

            Call CalculaProximoErro(sProxErro)

            Print #1, ""
            Print #1, "    'Informa tabela associada à Tela"
            Print #1, "    sTabela = " & """" & NomeArq.Text & """"
            Print #1, ""
            Print #1, "    'Lê os dados da Tela PedidoVenda"
            Print #1, "    lErro = Move_Tela_Memoria(" & sOBJ & ")"
            Print #1, "    If lErro <> SUCESSO Then gError " & sProxErro
            Print #1, ""
            Print #1, "    'Preenche a coleção colCampoValor, com nome do campo,"
            Print #1, "    'valor atual (com a tipagem do BD), tamanho do campo"
            Print #1, "    'no BD no caso de STRING e Key igual ao nome do campo"
            
            iIndice = 0
            For Each objColunasTabelas In colColunasTabelas
                iIndice = iIndice + 1
                With objColunasTabelas
                    If InStr(1, .sColuna, "NumInt") = 0 And InStr(1, .sColuna, "FilialEmpresa") = 0 And .iChave = MARCADO Then
                        
                        Call ObtemSiglaTipo(.sColunaTipo, sSigla, iTipo, sTipoVB)
                        
                        If sSigla = "s" Then
                            Print #1, "    colCampoValor.Add " & """" & .sColuna & """" & ", " & sOBJ & "." & .sAtributoClasse & ", UTILIZAR_STRING_TAMANHO_" & .lColunaTamanho & ", " & """" & .sColuna & """"
                        Else
                            Print #1, "    colCampoValor.Add " & """" & .sColuna & """" & ", " & sOBJ & "." & .sAtributoClasse & ", 0, " & """" & .sColuna & """"
                        End If
                    
                    End If
                    
                    If InStr(1, .sColuna, "FilialEmpresa") <> 0 And .iChave = MARCADO Then
                        Print #1, ""
                        Print #1, "    'Filtros para o Sistema de Setas"
                        Print #1, "    colSelecao.Add " & """" & "FilialEmpresa" & """" & ", OP_IGUAL, giFilialEmpresa"
                    End If
                
                End With
                
            Next
                    
            sScriptErro = sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro
        
        Case "Tela_Preenche"
        
            Print #1, ""
            
            iIndice = 0
            sAux = ""
            For Each objColunasTabelas In colColunasTabelas
                iIndice = iIndice + 1
                With objColunasTabelas
                    If InStr(1, .sColuna, "NumInt") = 0 And InStr(1, .sColuna, "FilialEmpresa") = 0 And .iChave = MARCADO Then
                        Print #1, TECLA_TAB & sOBJ & "." & .sAtributoClasse & " = colCampoValor.Item(" & """" & .sColuna & """" & ").vValor"
                    End If
                    
                    If InStr(1, .sColuna, "FilialEmpresa") <> 0 And .iChave = MARCADO Then
                        Print #1, ""
                        Print #1, TECLA_TAB & sOBJ & ".iFilialEmpresa = giFilialEmpresa"
                    End If
                    
                    Call ObtemSiglaTipo(.sColunaTipo, sSigla, iTipo, sTipoVB)

                    If .iChave = MARCADO Then
                    
                        If Len(Trim(sAux)) <> 0 Then sAux = sAux & " AND "
    
                        Select Case sSigla
                        
                            Case "s"
                                sAux = sAux & "Len(Trim(" & sOBJ & "." & .sAtributoClasse & ")) > 0"
                            
                            Case "dt"
                                sAux = sAux & sOBJ & "." & .sAtributoClasse & "<> DATA_NULA"
                            
                            Case "d", "l", "i"
                                sAux = sAux & sOBJ & "." & .sAtributoClasse & "<> 0"
                            
                            Case Else
                            
                        End Select
                        
                    End If
                
                End With
                
            Next
            
            Call CalculaProximoErro(sProxErro)
            
            Print #1, ""
            Print #1, TECLA_TAB & "If " & sAux & "Then"
            Print #1, ""
            Print #1, "        lErro = Traz_" & NomeArq.Text & "_Tela(" & sOBJ & ")"
            Print #1, "        If lErro <> SUCESSO Then gError " & sProxErro
            Print #1, ""
            Print #1, TECLA_TAB & "End If"
        
            sScriptErro = sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro
        
        Case "Gravar_Registro"
        
        
            Print #1, ""
            Print #1, "    GL_objMDIForm.MousePointer = vbHourglass"
            Print #1, ""
            Print #1, "    '#####################"
            Print #1, "    'CRITICA DADOS DA TELA"
            
            iIndice = 0
            For Each objColunasTabelas In colColunasTabelas
                            
                If objColunasTabelas.iChave = MARCADO Then
                
                    iIndice = iIndice + 1
                
                    If InStr(1, objColunasTabelas.sColuna, "NumInt") = 0 And InStr(1, objColunasTabelas.sColuna, "FilialEmpresa") = 0 Then
                    
                        Call CalculaProximoErro(sProxErro)
                        
                        Print #1, "    If Len(Trim(" & objColunasTabelas.sColuna & ".Text)) =0 then gError " & sProxErro
                    
                        sScriptErro = sScriptErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, <" & """" & "ERRO_" & UCase(objColunasTabelas.sColuna) & "_" & UCase(NomeArq.Text) & "_NAO_PREENCHIDO" & """" & ">, gErr)" & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & objColunasTabelas.sColuna & ".SetFocus" & sNL
                
                    End If
                    
                    If iIndice <> 1 Then
                        sChave = sChave & ", "
                    End If
                    
                    sChave = sChave & sOBJ & "." & objColunasTabelas.sAtributoClasse
                
                End If
            
            Next
            
            Call CalculaProximoErro(sProxErro)
            
            Print #1, "    '#####################"
            Print #1, ""
            Print #1, "    'Preenche o " & sOBJ
            Print #1, "    lErro = Move_Tela_Memoria(" & sOBJ & ")"
            Print #1, "    If lErro <> SUCESSO Then gError " & sProxErro
            
            sScriptErro = sScriptErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro
                        
            Call CalculaProximoErro(sProxErro)
                        
            Print #1, ""
            Print #1, "    lErro = Trata_Alteracao(" & sOBJ & ", " & sChave & ")"
            Print #1, "    If lErro <> SUCESSO Then gError " & sProxErro
            
            sScriptErro = sScriptErro & ", " & sProxErro
            
            Call CalculaProximoErro(sProxErro)
            
            Print #1, ""
            Print #1, "    'Grava o/a " & NomeArq.Text & " no Banco de Dados"
            Print #1, "    lErro = CF(""" & NomeArq.Text & "_Grava""" & ", " & sOBJ & ")"
            Print #1, "    If lErro <> SUCESSO Then gError " & sProxErro
            Print #1, ""
            Print #1, "    GL_objMDIForm.MousePointer = vbDefault"
        
            sScriptErro = sScriptErro & ", " & sProxErro
        
        Case "Limpa_Tela_" & NomeArq.Text
        
            Print #1, ""
            Print #1, "    'Fecha o comando das setas se estiver aberto"
            Print #1, "    Call ComandoSeta_Fechar(Me.Name)"
            Print #1, ""
            Print #1, "    'Função genérica que limpa campos da tela"
            Print #1, "    Call Limpa_Tela(Me)"
            Print #1, ""
            Print #1, "    iAlterado = 0"
        
        Case "Traz_" & NomeArq.Text & "_Tela"
        
            Print #1, "    Call Limpa_Tela_" & NomeArq.Text
            
            For Each objColunasTabelas In colColunasTabelas
            
                iIndice = iIndice + 1
                
                With objColunasTabelas
            
                    If InStr(1, .sColuna, "NumInt") = 0 And InStr(1, .sColuna, "FilialEmpresa") = 0 And .iChave = MARCADO Then
                    
                        Call ObtemSiglaTipo(.sColunaTipo, sSigla, iTipo, sTipoVB)
            
                        Select Case sSigla
                        
                            Case "s"
                                Print #1, TECLA_TAB & TECLA_TAB & .sColuna & ".text = " & sOBJ & "." & .sAtributoClasse
                            
                            Case "i"
                                Print #1, ""
                                Print #1, TECLA_TAB & TECLA_TAB & "If " & sOBJ & "." & .sAtributoClasse & " <> 0 Then "
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".PromptInclude = False "
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".text = Cstr(" & sOBJ & "." & .sAtributoClasse & ")"
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".PromptInclude = True "
                                Print #1, TECLA_TAB & TECLA_TAB & "End If"
                                Print #1, ""
                            
                            Case "dt"
                                Print #1, ""
                                Print #1, TECLA_TAB & TECLA_TAB & "If " & sOBJ & "." & .sAtributoClasse & " <> DATA_NULA Then "
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".PromptInclude = False "
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".text = Format(" & sOBJ & "." & .sAtributoClasse & "," & """" & "dd/mm/yy" & """" & ")"
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".PromptInclude = True "
                                Print #1, TECLA_TAB & TECLA_TAB & "End If"
                                Print #1, ""
                            
                            Case "d"
                                Print #1, ""
                                Print #1, TECLA_TAB & TECLA_TAB & "If " & sOBJ & "." & .sAtributoClasse & " <> 0 Then "
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".PromptInclude = False "
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".text = Format(" & sOBJ & "." & .sAtributoClasse & "," & .sColuna & ".Format)"
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".PromptInclude = True "
                                Print #1, TECLA_TAB & TECLA_TAB & "End If"
                                Print #1, ""
                    
                            Case "l"
                                Print #1, ""
                                Print #1, TECLA_TAB & TECLA_TAB & "If " & sOBJ & "." & .sAtributoClasse & " <> 0 Then "
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".PromptInclude = False "
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".text = Cstr(" & sOBJ & "." & .sAtributoClasse & ")"
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".PromptInclude = True "
                                Print #1, TECLA_TAB & TECLA_TAB & "End If"
                                Print #1, ""
                                
                        End Select
                        
                    End If
            
                End With
            
            Next
        
            Call CalculaProximoErro(sProxErro)
            
            Print #1, ""
            Print #1, "    'Lê o " & NomeArq.Text & " que está sendo Passado"
            Print #1, "    lErro = CF(""" & NomeArq.Text & "_Le" & """" & ", " & sOBJ & ")"
            Print #1, "    If lErro <> SUCESSO AND lErro <> " & sErroLeitura & " Then gError " & sProxErro
            Print #1, ""
            Print #1, "    If lErro = SUCESSO Then "
            Print #1, ""
            
            iIndice = 0
            For Each objColunasTabelas In colColunasTabelas
            
                iIndice = iIndice + 1
                
                With objColunasTabelas
        
                    If InStr(1, .sColuna, "NumInt") = 0 And InStr(1, .sColuna, "FilialEmpresa") = 0 Then
                    
                        Call ObtemSiglaTipo(.sColunaTipo, sSigla, iTipo, sTipoVB)
            
                        Select Case sSigla
                        
                            Case "s"
                                Print #1, TECLA_TAB & TECLA_TAB & .sColuna & ".text = " & sOBJ & "." & .sAtributoClasse
                            
                            Case "i"
                                Print #1, ""
                                Print #1, TECLA_TAB & TECLA_TAB & "If " & sOBJ & "." & .sAtributoClasse & " <> 0 Then "
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".PromptInclude = False "
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".text = Cstr(" & sOBJ & "." & .sAtributoClasse & ")"
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".PromptInclude = True "
                                Print #1, TECLA_TAB & TECLA_TAB & "End If"
                                Print #1, ""
                            
                            Case "dt"
                                Print #1, ""
                                Print #1, TECLA_TAB & TECLA_TAB & "If " & sOBJ & "." & .sAtributoClasse & " <> DATA_NULA Then "
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".PromptInclude = False "
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".text = Format(" & sOBJ & "." & .sAtributoClasse & "," & """" & "dd/mm/yy" & """" & ")"
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".PromptInclude = True "
                                Print #1, TECLA_TAB & TECLA_TAB & "End If"
                                Print #1, ""
                            
                            Case "d"
                                Print #1, ""
                                Print #1, TECLA_TAB & TECLA_TAB & "If " & sOBJ & "." & .sAtributoClasse & " <> 0 Then "
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".PromptInclude = False "
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".text = Format(" & sOBJ & "." & .sAtributoClasse & "," & .sColuna & ".Format)"
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".PromptInclude = True "
                                Print #1, TECLA_TAB & TECLA_TAB & "End If"
                                Print #1, ""
                    
                            Case "l"
                                Print #1, ""
                                Print #1, TECLA_TAB & TECLA_TAB & "If " & sOBJ & "." & .sAtributoClasse & " <> 0 Then "
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".PromptInclude = False "
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".text = Cstr(" & sOBJ & "." & .sAtributoClasse & ")"
                                Print #1, TECLA_TAB & TECLA_TAB & TECLA_TAB & .sColuna & ".PromptInclude = True "
                                Print #1, TECLA_TAB & TECLA_TAB & "End If"
                                Print #1, ""
                                
                        End Select
                        
                    End If
                
                End With
            
            Next
            
            Print #1, ""
            Print #1, "    End If "
            Print #1, ""
            Print #1, "    iAlterado = 0"
        
            sScriptErro = sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro
        
        Case "BotaoGravar_Click"
    
            Call CalculaProximoErro(sProxErro)
            
            Print #1, ""
            Print #1, "    lErro = Gravar_Registro"
            Print #1, "    If lErro <> SUCESSO Then gError " & sProxErro
            Print #1, ""
            Print #1, "    'Limpa Tela"
            Print #1, "    Call Limpa_Tela_" & NomeArq.Text
        
            sScriptErro = sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro
        
        Case "BotaoFechar_Click"
        
            Print #1, ""
            Print #1, "    Unload Me"
        
        Case "BotaoLimpar_Click"
        
            Call CalculaProximoErro(sProxErro)
            
            Print #1, ""
            Print #1, "    lErro = Teste_Salva(Me, iAlterado)"
            Print #1, "    If lErro <> SUCESSO Then gError " & sProxErro
            Print #1, ""
            Print #1, "    Call Limpa_Tela_" & NomeArq.Text
            
            sScriptErro = sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro
        
        Case "BotaoExcluir_Click"
        
            Print #1, ""
            Print #1, "    GL_objMDIForm.MousePointer = vbHourglass"
            Print #1, ""
            Print #1, "    '#####################"
            Print #1, "    'CRITICA DADOS DA TELA"
            
            iIndice = 0
            For Each objColunasTabelas In colColunasTabelas
            
                iIndice = iIndice + 1
                
                If objColunasTabelas.iChave = MARCADO Then
                
                    If InStr(1, objColunasTabelas.sColuna, "NumInt") = 0 And InStr(1, objColunasTabelas.sColuna, "FilialEmpresa") = 0 Then
                
                        Call CalculaProximoErro(sProxErro)
                        
                        Print #1, TECLA_TAB & "If Len(Trim(" & objColunasTabelas.sColuna & ".Text)) =0 then gError " & sProxErro
                    
                        sScriptErro = sScriptErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & "Call Rotina_Erro(vbOKOnly, <" & """" & "ERRO_" & UCase(objColunasTabelas.sColuna) & "_" & UCase(NomeArq.Text) & "_NAO_PREENCHIDO" & """" & ">, gErr)" & sNL & TECLA_TAB & TECLA_TAB & TECLA_TAB & objColunasTabelas.sColuna & ".SetFocus" & sNL
                
                    End If
                
                End If
            
            Next
            
            Print #1, "    '#####################"
            Print #1, ""
                    
            sAux = 0
            iIndice = 0
            For Each objColunasTabelas In colColunasTabelas
            
                iIndice = iIndice + 1
                
                With objColunasTabelas
                
                    If InStr(1, .sColuna, "NumInt") = 0 And InStr(1, .sColuna, "FilialEmpresa") = 0 And .iChave = MARCADO Then
        
                        Call ObtemSiglaTipo(.sColunaTipo, sSigla, iTipo, sTipoVB)
            
                        Select Case sSigla
                        
                            Case "s"
                                Print #1, TECLA_TAB & sOBJ & "." & .sAtributoClasse & " = " & .sColuna & ".text"
                            
                            Case "i"
                                Print #1, TECLA_TAB & sOBJ & "." & .sAtributoClasse & " = StrParaInt(" & .sColuna & ".text)"
                            
                            Case "dt"
                                Print #1, TECLA_TAB & sOBJ & "." & .sAtributoClasse & " = StrParaDate(" & .sColuna & ".text)"
                            
                            Case "d"
                                Print #1, TECLA_TAB & sOBJ & "." & .sAtributoClasse & " = StrParaDbl(" & .sColuna & ".text)"
                    
                            Case "l"
                                Print #1, TECLA_TAB & sOBJ & "." & .sAtributoClasse & " = StrParaLong(" & .sColuna & ".text)"
                    
                        End Select
                        
                        sAux = sOBJ & "." & .sAtributoClasse
                        
                    End If
                    
                    If InStr(1, .sColuna, "FilialEmpresa") <> 0 And .iChave = MARCADO Then
                        Print #1, TECLA_TAB & sOBJ & "." & .sAtributoClasse & " = giFilialEmpresa"
                    End If
                
                End With
            
            Next
            Print #1, ""
            Print #1, TECLA_TAB & "'Pergunta ao usuário se confirma a exclusão"
            Print #1, TECLA_TAB & "vbMsgRes = Rotina_Aviso(vbYesNo, " & """" & "AVISO_CONFIRMA_EXCLUSAO_" & UCase(NomeArq.Text) & """" & ", " & sAux & ")"
            Print #1, ""
            Print #1, TECLA_TAB & "If vbMsgRes = vbYes Then"
            Print #1, ""
                        
            Call CalculaProximoErro(sProxErro)
            
            Print #1, TECLA_TAB & "    'Exclui a requisição de consumo"
            Print #1, TECLA_TAB & "    lErro = CF(" & """" & NomeArq.Text & "_Exclui" & """" & ", " & sOBJ & ")"
            Print #1, TECLA_TAB & "    If lErro <> SUCESSO Then gError " & sProxErro
            Print #1, ""
            Print #1, TECLA_TAB & "    'Limpa Tela"
            Print #1, TECLA_TAB & "    Call Limpa_Tela_" & NomeArq.Text
            Print #1, ""
            Print #1, TECLA_TAB & "End If"
            Print #1, ""
            Print #1, "    GL_objMDIForm.MousePointer = vbDefault"
        
            sScriptErro = sScriptErro & sNL & TECLA_TAB & TECLA_TAB & "Case " & sProxErro
        
        Case Else
        
    End Select
    
    Comandos_Rotina = SUCESSO

    Exit Function

Erro_Comandos_Rotina:

    Comandos_Rotina = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143970)

    End Select
    
    Exit Function
    
End Function

Private Function Erros_Rotina(ByVal sNomeRot As String, ByVal sScriptErro As String) As Long

Dim lErro As Long

On Error GoTo Erro_Erros_Rotina

    If Len(Trim(sScriptErro)) > 0 Then Print #1, sScriptErro
    
    Erros_Rotina = SUCESSO

    Exit Function

Erro_Erros_Rotina:

    Erros_Rotina = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143971)

    End Select
    
    Exit Function
    
End Function

Private Function Final1_Rotina(ByVal sNomeRot As String) As Long

Dim lErro As Long

On Error GoTo Erro_Final1_Rotina

    Select Case sNomeRot
    
        Case "Trata_Parametros"
        
        Case "Move_Tela_Memoria"
        
        Case "Tela_Extrai"
        
        Case "Tela_Preenche"
        
        Case "Gravar_Registro"
            Print #1, ""
            Print #1, "    GL_objMDIForm.MousePointer = vbDefault"
        
        Case "Limpa_Tela_" & NomeArq.Text
        
        Case "Form_Load"
            Print #1, ""
            Print #1, "    lErro_Chama_Tela = gErr"
        
        Case "BotaoGravar_Click"
        
        Case "BotaoFechar_Click"
        
        Case "BotaoLimpar_Click"
        
        Case "BotaoExcluir_Click"
            Print #1, ""
            Print #1, "    GL_objMDIForm.MousePointer = vbDefault"
        
        Case "Traz_" & NomeArq.Text & "_Tela"
        
        Case Else
        
    End Select
    
    Final1_Rotina = SUCESSO

    Exit Function

Erro_Final1_Rotina:

    Final1_Rotina = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143972)

    End Select
    
    Exit Function
    
End Function

Private Function Final2_Rotina(ByVal sNomeRot As String) As Long

Dim lErro As Long

On Error GoTo Erro_Final2_Rotina

    Select Case sNomeRot
    
        Case "Trata_Parametros"
            Print #1, ""
            Print #1, "    iAlterado = 0"
        
        Case "Move_Tela_Memoria"
        
        Case "Tela_Extrai"
        
        Case "Tela_Preenche"
        
        Case "Gravar_Registro"
        
        Case "Limpa_Tela_" & NomeArq.Text
        
        Case "BotaoGravar_Click"
        
        Case "BotaoFechar_Click"
        
        Case "BotaoLimpar_Click"
        
        Case "BotaoExcluir_Click"
        
        Case "Form_Load"
            Print #1, ""
            Print #1, "    iAlterado = 0"
        
        Case "Traz_" & NomeArq.Text & "_Tela"
        
        Case Else
        
    End Select
    
    Final2_Rotina = SUCESSO

    Exit Function

Erro_Final2_Rotina:

    Final2_Rotina = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143973)

    End Select
    
    Exit Function
    
End Function

Private Function CTL_Cria_Inicial(colCampo As Collection) As Long

Dim lErro As Long

On Error GoTo Erro_CTL_Cria_Inicial

    Print #1, "Version 5.0"
    Print #1, "Object = " & """" & "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0" & """" & "; " & """" & "MSMASK32.OCX" & """"
    Print #1, "Object = " & """" & "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0" & """" & "; " & """" & "mscomctl.OCX" & """"
    Print #1, "Object = " & """" & "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0" & """" & "; " & """" & "MSCOMCT2.OCX" & """"
    Print #1, "Begin VB.UserControl " & NomeTela.Text
    Print #1, "   ClientHeight = 6000"
    Print #1, "   ClientLeft = 0"
    Print #1, "   ClientTop = 0"
    Print #1, "   ClientWidth = 9510"
    Print #1, "   KeyPreview = -1         'True"
    Print #1, "   ScaleHeight = 5745"
    Print #1, "   ScaleWidth = 8145"
    Print #1, "   Begin VB.PictureBox Picture1"
    Print #1, "      Height = 510"
    Print #1, "      Left = 7320"
    Print #1, "      ScaleHeight = 450"
    Print #1, "      ScaleWidth = 2025"
    Print #1, "      TabIndex = 0"
    Print #1, "      TabStop = 0             'False"
    Print #1, "      Top = 30"
    Print #1, "      Width = 2085"
    Print #1, "      Begin VB.CommandButton BotaoGravar"
    Print #1, "         Height = 360"
    Print #1, "         Left = 60"
    Print #1, "         Picture         =   """ & NomeTela.Text & ".ctx" & """:0000"
    Print #1, "         Style = 1              'Graphical"
    Print #1, "         TabIndex = 1"
    Print #1, "         ToolTipText = " & """" & "Gravar" & """"
    Print #1, "         Top = 45"
    Print #1, "         Width = 420"
    Print #1, "      End"
    Print #1, "      Begin VB.CommandButton BotaoExcluir"
    Print #1, "         Height = 360"
    Print #1, "         Left = 570"
    Print #1, "         Picture         =   """ & NomeTela.Text & ".ctx" & """:015A"
    Print #1, "         Style = 1              'Graphical"
    Print #1, "         TabIndex = 2"
    Print #1, "         ToolTipText = """ & "Excluir" & """"
    Print #1, "         Top = 45"
    Print #1, "         Width = 420"
    Print #1, "      End"
    Print #1, "      Begin VB.CommandButton BotaoLimpar"
    Print #1, "         Height = 360"
    Print #1, "         Left = 1065"
    Print #1, "         Picture         =   """ & NomeTela.Text & ".ctx" & """:02E4"
    Print #1, "         Style = 1              'Graphical"
    Print #1, "         TabIndex = 3"
    Print #1, "         ToolTipText = """ & "Limpar" & """"
    Print #1, "         Top = 45"
    Print #1, "         Width = 420"
    Print #1, "      End"
    Print #1, "      Begin VB.CommandButton BotaoFechar"
    Print #1, "         Height = 360"
    Print #1, "         Left = 1545"
    Print #1, "         Picture         =   """ & NomeTela.Text & ".ctx" & """:0816"
    Print #1, "         Style = 1              'Graphical"
    Print #1, "         TabIndex = 4"
    Print #1, "         ToolTipText = """ & "Fechar" & """"
    Print #1, "         Top = 45"
    Print #1, "         Width = 420"
    Print #1, "      End"
    Print #1, "   End"
    
    lErro = Controles_Tela_Cria(colCampo)
    If lErro <> SUCESSO Then gError 131799
    
    Print #1, "End"
    Print #1, "Attribute VB_Name = """ & NomeTela.Text & """"
    Print #1, "Attribute VB_GlobalNameSpace = False"
    Print #1, "Attribute VB_Creatable = True"
    Print #1, "Attribute VB_PredeclaredId = False"
    Print #1, "Attribute VB_Exposed = True"
    Print #1, "Option Explicit"
    
    CTL_Cria_Inicial = SUCESSO

    Exit Function

Erro_CTL_Cria_Inicial:

    CTL_Cria_Inicial = gErr

    Select Case gErr
    
        Case 131799

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143974)

    End Select
    
    Exit Function
    
End Function
'FIM DOS GERADORES DE SCRIPTS
'####################################################

'####################################################
'CRITICAS

Private Function Critica_Tela() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objColunasTabelas As ClassColunasTabelas
Dim iPos As Integer
Dim iContarB As Integer
Dim iContarI As Integer
Dim vbResult As VbMsgBoxResult
Dim bExisteClasse As Boolean
Dim sClasse  As String
Dim sAtributo As String

On Error GoTo Erro_Critica_Tela

    If Len(Trim(DescBrowse.Text)) = 0 Then gError 131659
    If Len(Trim(NomeArq.Text)) = 0 Then gError 131660
    If Len(Trim(NomeBrowse.Text)) = 0 Then gError 131661
    If Len(Trim(Classe.Text)) = 0 Then gError 131662
    If Len(Trim(NomeTela.Text)) = 0 Then gError 131663
    If Len(Trim(ModuloAcesso.Text)) = 0 Then gError 131664
    If Len(Trim(ModuloFormata.Text)) = 0 Then gError 131664
    If Len(Trim(ModuloTela.Text)) = 0 Then gError 131664
    If Len(Trim(ModuloClasse.Text)) = 0 Then gError 131664

    iContarB = 0
    iContarI = 0
    bExisteClasse = False
    
    'Se não vai fazer a Classe, verifica se ele já existe
    If optClasse.Value <> vbChecked Then
    
        sClasse = "Globais" & ModuloClasse.Text & "." & Classe.Text
    
        lErro = Critica_Objeto(sClasse)
        If lErro <> SUCESSO Then
        
            bExisteClasse = False
            
            vbResult = Rotina_Aviso(vbYesNo, "AVISO_CLASSE_INEXISTENTE")
    
            If vbResult = vbNo Then gError 131746
       
        Else
        
            bExisteClasse = True
       
        End If
        
    End If

    For iIndice = 1 To objGridColunas.iLinhasExistentes
    
        iPos = iIndice

        If StrParaInt(GridColunas.TextMatrix(iIndice, iGrid_Browse_Col)) = MARCADO Then
            iContarB = iContarB + 1
        End If
        
        If StrParaInt(GridColunas.TextMatrix(iIndice, iGrid_Indice_Col)) = MARCADO Then
            iContarI = iContarI + 1
        End If
               
        If Len(Trim(GridColunas.TextMatrix(iIndice, iGrid_Coluna_Col))) = 0 Then gError 131665
        If StrParaInt(GridColunas.TextMatrix(iIndice, iGrid_Ordem_Col)) = 0 Then gError 131666
        If StrParaLong(GridColunas.TextMatrix(iIndice, iGrid_Tamanho_Col)) = 0 Then gError 131668
        If Len(Trim(GridColunas.TextMatrix(iIndice, iGrid_Tipo_Col))) = 0 Then gError 131669
        
        If StrParaInt(GridColunas.TextMatrix(iIndice, iGrid_TemClasse_Col)) = MARCADO Then
            sAtributo = GridColunas.TextMatrix(iIndice, iGrid_AtribClasse_Col)
            If Len(Trim(sAtributo)) = 0 Then gError 131670
        End If
        
        If StrParaLong(GridColunas.TextMatrix(iIndice, iGrid_TamanhoTela_Col)) = 0 Then gError 131671
        If Len(Trim(GridColunas.TextMatrix(iIndice, iGrid_Descricao_Col))) = 0 Then gError 131672
        
        If (bExisteClasse) And (StrParaInt(GridColunas.TextMatrix(iIndice, iGrid_TemClasse_Col)) = MARCADO) Then
            lErro = Critica_ObjetoAtributo(sClasse, sAtributo)
            If lErro <> SUCESSO Then gError 131920
        End If
        
    Next
    
    If iContarB = 0 Then
        
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_SEM_BROWSE_MARCADO")
    
        If vbResult = vbNo Then gError 131744
    
    End If
    
    If iContarI = 0 Then
        
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_SEM_INDICE_MARCADO")

        If vbResult = vbNo Then gError 131745

    End If
        
    Critica_Tela = SUCESSO

    Exit Function

Erro_Critica_Tela:

    Critica_Tela = gErr

    Select Case gErr
    
        Case 131659
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)
            Descricao.SetFocus
        
        Case 131660
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_PREENCHIDO", gErr)
            NomeArq.SetFocus
        
        Case 131661
            Call Rotina_Erro(vbOKOnly, "ERRO_BROWSE_NAO_PREENCHIDO", gErr)
            NomeBrowse.SetFocus
        
        Case 131662
            Call Rotina_Erro(vbOKOnly, "ERRO_CLASSE_NAO_PREENCHIDO", gErr)
            Classe.SetFocus
        
        Case 131663
            Call Rotina_Erro(vbOKOnly, "ERRO_TELA_NAO_INFORMADA", gErr)
            NomeTela.SetFocus
        
        Case 131664
            Call Rotina_Erro(vbOKOnly, "ERRO_MODULO_NAO_PREENCHIDO", gErr)

        Case 131665
            Call Rotina_Erro(vbOKOnly, "ERRO_COLUNA_GRID_NAO_PREENCHIDO", gErr, iPos)

        Case 131666
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEM_GRID_NAO_PREENCHIDO", gErr, iPos)

        Case 131667
            Call Rotina_Erro(vbOKOnly, "ERRO_PRECISAO_GRID_NAO_PREENCHIDO", gErr, iPos)

        Case 131668
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_GRID_NAO_PREENCHIDO", gErr, iPos)

        Case 131669
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_GRID_NAO_PREENCHIDO", gErr, iPos)

        Case 131670
            Call Rotina_Erro(vbOKOnly, "ERRO_ATRIBCLASSE_GRID_NAO_PREENCHIDO", gErr, iPos)

        Case 131671
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHOTELA_GRID_NAO_PREENCHIDO", gErr, iPos)

        Case 131672
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_GRID_NAO_PREENCHIDO", gErr, iPos)

        Case 131744, 131745, 131746
        
        Case 131920
            Call Rotina_Erro(vbOKOnly, "ERRO_ATRIBUTO_GRID_NAO_EXISTE", gErr, sAtributo, iPos, sClasse)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143975)

    End Select

    Exit Function

End Function

Private Function Critica_Objeto(ByVal sOBJ As String) As Long

Dim objClasse As Object

On Error GoTo Erro_Critica_Objeto

    Set objClasse = CreateObject(sOBJ)

    Critica_Objeto = SUCESSO

    Exit Function

Erro_Critica_Objeto:

    Critica_Objeto = gErr
    
    Select Case gErr

        Case Else

    End Select

    Exit Function

End Function

Private Function Critica_ObjetoAtributo(ByVal sOBJ As String, ByVal sAtributo As String) As Long

Dim objClasse As Object

On Error GoTo Erro_Critica_ObjetoAtributo

    Set objClasse = CreateObject(sOBJ)

    Call CallByName(objClasse, sAtributo, VbLet, 0)

    Critica_ObjetoAtributo = SUCESSO

    Exit Function

Erro_Critica_ObjetoAtributo:

    Critica_ObjetoAtributo = gErr
    
    Select Case gErr

        Case Else

    End Select

    Exit Function

End Function
'FIM DAS CRITICAS
'###################################################

'###################################################
'SCRIPT DO GRID
Private Sub GridColunas_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentes As Integer
Dim iLinhaAtual As Integer

    iLinhasExistentes = objGridColunas.iLinhasExistentes
    iLinhaAtual = GridColunas.Row

    Call Grid_Trata_Tecla1(KeyCode, objGridColunas)

End Sub

Private Sub GridColunas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridColunas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridColunas, iAlterado)
    End If

End Sub

Private Sub GridColunas_EnterCell()

    Call Grid_Entrada_Celula(objGridColunas, iAlterado)

End Sub

Private Sub GridColunas_GotFocus()

    Call Grid_Recebe_Foco(objGridColunas)

End Sub

Private Sub GridColunas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridColunas, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridColunas, iAlterado)
    End If

End Sub

Private Sub GridColunas_LeaveCell()

    Call Saida_Celula(objGridColunas)

End Sub

Private Sub GridColunas_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridColunas)
    
End Sub

Private Sub GridColunas_RowColChange()

    Call Grid_RowColChange(objGridColunas)

End Sub

Private Sub GridColunas_Scroll()

    Call Grid_Scroll(objGridColunas)

End Sub
'FIM DO SCRIPT DO GRID
'#########################################################

'#########################################################
'SCRIPT DE CAMPOS DE GRID
Private Sub Coluna_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridColunas)

End Sub

Private Sub Coluna_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridColunas)

End Sub

Private Sub Coluna_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridColunas.objControle = Coluna
    lErro = Grid_Campo_Libera_Foco(objGridColunas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Indice_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridColunas)

End Sub

Private Sub Indice_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridColunas)

End Sub

Private Sub Indice_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridColunas.objControle = Indice
    lErro = Grid_Campo_Libera_Foco(objGridColunas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Browse_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridColunas)

End Sub

Private Sub Browse_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridColunas)

End Sub

Private Sub Browse_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridColunas.objControle = Browse
    lErro = Grid_Campo_Libera_Foco(objGridColunas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub optBrowse_Click()

    If optBrowse.Value = False Then optTodos.Value = False

End Sub

Private Sub optClasse_Click()

    If optClasse.Value = False Then optTodos.Value = False

End Sub

Private Sub optDic_Click()

    If optDic.Value = False Then optTodos.Value = False

End Sub

Private Sub optExclusao_Click()

    If optExclusao.Value = False Then optTodos.Value = False

End Sub

Private Sub optGravacao_Click()

    If optGravacao.Value = False Then optTodos.Value = False

End Sub

Private Sub optLeitura_Click()

    If optLeitura.Value = False Then optTodos.Value = False

End Sub

Private Sub optTela_Click()

    If optTela.Value = False Then optTodos.Value = False

End Sub

Private Sub optTodos_Click()

    If optTodos.Value = vbChecked Then
        optBrowse.Value = vbChecked
        optClasse.Value = vbChecked
        optDic.Value = vbChecked
        optExclusao.Value = vbChecked
        optGravacao.Value = vbChecked
        optLeitura.Value = vbChecked
        optTela.Value = vbChecked
        optBrowse.Value = vbChecked
        optType.Value = vbChecked
    
        optBrowse.Enabled = False
        optClasse.Enabled = False
        optDic.Enabled = False
        optExclusao.Enabled = False
        optGravacao.Enabled = False
        optLeitura.Enabled = False
        optTela.Enabled = False
        optBrowse.Enabled = False
        optType.Enabled = False
    
    Else
        
        optBrowse.Enabled = True
        optClasse.Enabled = True
        optDic.Enabled = True
        optExclusao.Enabled = True
        optGravacao.Enabled = True
        optLeitura.Enabled = True
        optTela.Enabled = True
        optBrowse.Enabled = True
        optType.Enabled = True
        
    End If

End Sub

Private Sub optType_Click()

    If optType.Value = False Then optTodos.Value = False

End Sub

Private Sub Tipo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridColunas)

End Sub

Private Sub Tipo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridColunas)

End Sub

Private Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridColunas.objControle = Tipo
    lErro = Grid_Campo_Libera_Foco(objGridColunas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Tamanho_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridColunas)

End Sub

Private Sub Tamanho_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridColunas)

End Sub

Private Sub Tamanho_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridColunas.objControle = Tamanho
    lErro = Grid_Campo_Libera_Foco(objGridColunas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TamanhoTela_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridColunas)

End Sub

Private Sub TamanhoTela_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridColunas)

End Sub

Private Sub TamanhoTela_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridColunas.objControle = TamanhoTela
    lErro = Grid_Campo_Libera_Foco(objGridColunas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Precisao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridColunas)

End Sub

Private Sub Precisao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridColunas)

End Sub

Private Sub Precisao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridColunas.objControle = Precisao
    lErro = Grid_Campo_Libera_Foco(objGridColunas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Ordem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridColunas)

End Sub

Private Sub Ordem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridColunas)

End Sub

Private Sub Ordem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridColunas.objControle = Ordem
    lErro = Grid_Campo_Libera_Foco(objGridColunas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Descricao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridColunas)

End Sub

Private Sub Descricao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridColunas)

End Sub

Private Sub Descricao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridColunas.objControle = Descricao
    lErro = Grid_Campo_Libera_Foco(objGridColunas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TemClasse_click()
    
Dim sSigla As String
Dim sTipo As String
Dim sNome As String
Dim iTipo As Integer
Dim sTipoVB As String

    If GridColunas.TextMatrix(GridColunas.Row, iGrid_TemClasse_Col) = vbChecked Then
    
        sTipo = GridColunas.TextMatrix(GridColunas.Row, iGrid_Tipo_Col)
        sNome = GridColunas.TextMatrix(GridColunas.Row, iGrid_Coluna_Col)
    
        Call ObtemSiglaTipo(sTipo, sSigla, iTipo, sTipoVB)
        GridColunas.TextMatrix(GridColunas.Row, iGrid_AtribClasse_Col) = sSigla & sNome
    
    Else
        GridColunas.TextMatrix(GridColunas.Row, iGrid_AtribClasse_Col) = ""
    
    End If

End Sub

Private Sub TemClasse_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridColunas)

End Sub

Private Sub TemClasse_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridColunas)

End Sub

Private Sub TemClasse_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridColunas.objControle = TemClasse
    lErro = Grid_Campo_Libera_Foco(objGridColunas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Chave_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridColunas)

End Sub

Private Sub Chave_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridColunas)

End Sub

Private Sub Chave_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridColunas.objControle = Chave
    lErro = Grid_Campo_Libera_Foco(objGridColunas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub SubTipo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridColunas)

End Sub

Private Sub SubTipo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridColunas)

End Sub

Private Sub SubTipo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridColunas.objControle = SubTipo
    lErro = Grid_Campo_Libera_Foco(objGridColunas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub AtribClasse_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridColunas)

End Sub

Private Sub AtribClasse_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridColunas)

End Sub

Private Sub AtribClasse_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridColunas.objControle = AtribClasse
    lErro = Grid_Campo_Libera_Foco(objGridColunas)
    If lErro <> SUCESSO Then Cancel = True

End Sub
'FIM DOS SCRIPTS DE CAMPOS DO GRID
'######################################################

'########################################################
'SAIDA DE CELULA
Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual a coluna do Grid em questão
        Select Case objGridInt.objGrid.Col
            
            Case iGrid_AtribClasse_Col
                lErro = Saida_Celula_AtribClasse(objGridInt)
                If lErro <> SUCESSO Then gError 39355
                
            Case iGrid_Browse_Col
                lErro = Saida_Celula_Browse(objGridInt)
                If lErro <> SUCESSO Then gError 39355
                
            Case iGrid_Coluna_Col
                lErro = Saida_Celula_Coluna(objGridInt)
                If lErro <> SUCESSO Then gError 39355
                
            Case iGrid_Ordem_Col
                lErro = Saida_Celula_Ordem(objGridInt)
                If lErro <> SUCESSO Then gError 39355
                
            Case iGrid_Precisao_Col
                lErro = Saida_Celula_Precisao(objGridInt)
                If lErro <> SUCESSO Then gError 39355
                
            Case iGrid_Tamanho_Col
                lErro = Saida_Celula_Tamanho(objGridInt)
                If lErro <> SUCESSO Then gError 39355
                
            Case iGrid_TamanhoTela_Col
                lErro = Saida_Celula_TamanhoTela(objGridInt)
                If lErro <> SUCESSO Then gError 39355
        
            Case iGrid_Tipo_Col
                lErro = Saida_Celula_Tipo(objGridInt)
                If lErro <> SUCESSO Then gError 39355
        
            Case iGrid_TemClasse_Col
                lErro = Saida_Celula_TemClasse(objGridInt)
                If lErro <> SUCESSO Then gError 39355
        
            Case iGrid_Descricao_Col
                lErro = Saida_Celula_Descricao(objGridInt)
                If lErro <> SUCESSO Then gError 39355
        
            Case iGrid_Indice_Col
                lErro = Saida_Celula_Indice(objGridInt)
                If lErro <> SUCESSO Then gError 39355
        
            Case iGrid_Chave_Col
                lErro = Saida_Celula_Chave(objGridInt)
                If lErro <> SUCESSO Then gError 39355
        
            Case iGrid_SubTipo_Col
                lErro = Saida_Celula_SubTipo(objGridInt)
                If lErro <> SUCESSO Then gError 39355
        
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 39356

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 39355

        Case 39356
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143976)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Coluna(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Versão do Kit que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Coluna

    Set objGridInt.objControle = Coluna

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129217

    Saida_Celula_Coluna = SUCESSO

    Exit Function

Erro_Saida_Celula_Coluna:

    Saida_Celula_Coluna = gErr

    Select Case gErr

        Case 129217
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143977)

    End Select
    
    Exit Function

End Function

Private Function Saida_Celula_TamanhoTela(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Versão do Kit que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_TamanhoTela

    Set objGridInt.objControle = TamanhoTela
    
    If Len(Trim(TamanhoTela.Text)) > 0 Then
     
        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(TamanhoTela.Text)
        If lErro <> SUCESSO Then gError 131739
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129217

    Saida_Celula_TamanhoTela = SUCESSO

    Exit Function

Erro_Saida_Celula_TamanhoTela:

    Saida_Celula_TamanhoTela = gErr

    Select Case gErr

        Case 129217, 131739
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143978)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_AtribClasse(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Versão do Kit que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_AtribClasse

    Set objGridInt.objControle = AtribClasse

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129217

    Saida_Celula_AtribClasse = SUCESSO

    Exit Function

Erro_Saida_Celula_AtribClasse:

    Saida_Celula_AtribClasse = gErr

    Select Case gErr

        Case 129217
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143979)

    End Select
    
    Exit Function

End Function

Private Function Saida_Celula_Tipo(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Versão do Kit que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Tipo

    Set objGridInt.objControle = Tipo

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129217

    Saida_Celula_Tipo = SUCESSO

    Exit Function

Erro_Saida_Celula_Tipo:

    Saida_Celula_Tipo = gErr

    Select Case gErr

        Case 129217
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143980)

    End Select
    
    Exit Function

End Function

Private Function Saida_Celula_Tamanho(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Versão do Kit que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Tamanho

    Set objGridInt.objControle = Tamanho

    If Len(Trim(Tamanho.Text)) > 0 Then
     
        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(Tamanho.Text)
        If lErro <> SUCESSO Then gError 131739
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129217

    Saida_Celula_Tamanho = SUCESSO

    Exit Function

Erro_Saida_Celula_Tamanho:

    Saida_Celula_Tamanho = gErr

    Select Case gErr

        Case 129217, 131739
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143981)

    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_Browse(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Versão do Kit que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Browse

    Set objGridInt.objControle = Browse
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129217

    Saida_Celula_Browse = SUCESSO

    Exit Function

Erro_Saida_Celula_Browse:

    Saida_Celula_Browse = gErr

    Select Case gErr

        Case 129217
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143982)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Indice(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Versão do Kit que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Indice

    Set objGridInt.objControle = Indice
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129217

    Saida_Celula_Indice = SUCESSO

    Exit Function

Erro_Saida_Celula_Indice:

    Saida_Celula_Indice = gErr

    Select Case gErr

        Case 129217
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143983)

    End Select
    
    Exit Function

End Function

Private Function Saida_Celula_Chave(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Versão do Kit que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Chave

    Set objGridInt.objControle = Chave
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129217

    Saida_Celula_Chave = SUCESSO

    Exit Function

Erro_Saida_Celula_Chave:

    Saida_Celula_Chave = gErr

    Select Case gErr

        Case 129217
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143984)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Precisao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Versão do Kit que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Precisao

    Set objGridInt.objControle = Precisao

    If Len(Trim(Precisao.Text)) > 0 Then
     
        'Critica se valor é positivo
        lErro = Valor_NaoNegativo_Critica(Precisao.Text)
        If lErro <> SUCESSO Then gError 131739
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129217

    Saida_Celula_Precisao = SUCESSO

    Exit Function

Erro_Saida_Celula_Precisao:

    Saida_Celula_Precisao = gErr

    Select Case gErr

        Case 129217, 131739
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143985)

    End Select
    
    Exit Function

End Function

Private Function Saida_Celula_SubTipo(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Sub Tipo que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_SubTipo

    Set objGridInt.objControle = SubTipo

    If Len(Trim(SubTipo.Text)) > 0 Then
     
        'Critica se valor é positivo
        lErro = Valor_NaoNegativo_Critica(SubTipo.Text)
        If lErro <> SUCESSO Then gError 131739
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129217

    Saida_Celula_SubTipo = SUCESSO

    Exit Function

Erro_Saida_Celula_SubTipo:

    Saida_Celula_SubTipo = gErr

    Select Case gErr

        Case 129217, 131739
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143986)

    End Select
    
    Exit Function

End Function

Private Function Saida_Celula_Ordem(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Versão do Kit que está deixando de ser a corrente

Dim lErro As Long
Dim iValorNovo As Integer
Dim iValorAntigo As Integer
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_Ordem

    Set objGridInt.objControle = Ordem
    
    If Len(Trim(Ordem.Text)) > 0 Then
     
        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(Ordem.Text)
        If lErro <> SUCESSO Then gError 131739
        
        iLinha = GridColunas.Row
        iValorAntigo = StrParaInt(GridColunas.TextMatrix(iLinha, iGrid_Ordem_Col))
        iValorNovo = StrParaInt(Ordem.Text)
        
        If iValorNovo > objGridColunas.iLinhasExistentes Then gError 131740
        
        Call MantemSequencialOrdem(iValorAntigo, iValorNovo, iLinha)
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129217
    
    Saida_Celula_Ordem = SUCESSO

    Exit Function

Erro_Saida_Celula_Ordem:

    Saida_Celula_Ordem = gErr

    Select Case gErr
    
        Case 131740
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_ORD_INVALIDO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 129217, 131739
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143987)

    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_TemClasse(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Versão do Kit que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_TemClasse

    Set objGridInt.objControle = TemClasse

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129217

    Saida_Celula_TemClasse = SUCESSO

    Exit Function

Erro_Saida_Celula_TemClasse:

    Saida_Celula_TemClasse = gErr

    Select Case gErr

        Case 129217
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143988)

    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_Descricao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Versão do Kit que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Descricao

    Set objGridInt.objControle = Descricao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129217

    Saida_Celula_Descricao = SUCESSO

    Exit Function

Erro_Saida_Celula_Descricao:

    Saida_Celula_Descricao = gErr

    Select Case gErr

        Case 129217
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143989)

    End Select
    
    Exit Function

End Function
'FIM DA SAIDA DE CELULA
'########################################################

'########################################################
'MARCA E DESMARCA CAMPOS DO GRID
Private Sub DesmarcarBrowse_Click()
    Call Marca_Desmarca(iGrid_Browse_Col, DESMARCADO)
End Sub

Private Sub DesmarcarClasses_Click()
    Call Marca_Desmarca(iGrid_TemClasse_Col, DESMARCADO)
End Sub

Private Sub DesmarcarIndices_Click()
    Call Marca_Desmarca(iGrid_Indice_Col, DESMARCADO)
End Sub

Private Sub MarcarBrowse_Click()
    Call Marca_Desmarca(iGrid_Browse_Col, MARCADO)
End Sub

Private Sub MarcarClasse_Click()
    Call Marca_Desmarca(iGrid_TemClasse_Col, MARCADO)
End Sub

Private Sub MarcarIndices_Click()
    Call Marca_Desmarca(iGrid_Indice_Col, MARCADO)
End Sub

Private Sub MarcaChave_Click()
    Call Marca_Desmarca(iGrid_Chave_Col, MARCADO)
End Sub

Private Sub DesmarcaChave_Click()
    Call Marca_Desmarca(iGrid_Chave_Col, DESMARCADO)
End Sub
'FIM DE MARCA E DESCAMAR CAMPOS DO GRID
'########################################################

'#############################################
'INSERIDO POR WAGNER
'ROTINAS DE BD
Public Function ColunasTabelas_Le(ByVal sNomeArq As String, ByVal colColunasTabelas As Collection) As Long
'Lê syscolumns e sysobjects

Dim lErro As Long
Dim lComando As Long
Dim tColunasTabelas As typeColunasTabelas
Dim objColunasTabelas As ClassColunasTabelas

On Error GoTo Erro_ColunasTabelas_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 131700
    
    With tColunasTabelas
    
        'Aloca espaço no buffer
        .sArquivo = String(STRING_STRING_MAX, 0)
        .sArquivoTipo = String(STRING_STRING_MAX, 0)
        .sColuna = String(STRING_STRING_MAX, 0)
        .sColunaTipo = String(STRING_STRING_MAX, 0)
    
        'Le o syscolumns e sysobjects
        lErro = Comando_Executar(lComando, "SELECT O.Name,O.xtype, C.Name, T.Name, C.length, CONVERT(int,C.xprec) " & _
                                            "FROM syscolumns AS C, sysobjects AS O, systypes AS T " & _
                                            "WHERE O.id = C.id AND C.xtype = T.xtype AND O.name = ? ORDER BY C.colorder ", _
                                            .sArquivo, .sArquivoTipo, .sColuna, .sColunaTipo, .lColunaTamanho, .lColunaPrecisao, sNomeArq)
        If lErro <> AD_SQL_SUCESSO Then gError 131701
    
    End With

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131702

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objColunasTabelas = New ClassColunasTabelas

        With objColunasTabelas
        
            .sArquivo = Trim(tColunasTabelas.sArquivo)
            .sArquivoTipo = Trim(tColunasTabelas.sArquivoTipo)
            .sColuna = Trim(tColunasTabelas.sColuna)
            .sColunaTipo = Trim(tColunasTabelas.sColunaTipo)
            If .sArquivoTipo = ARQUIVO_VIEW And .sColunaTipo = "string" Then
                .lColunaTamanho = 255
                .lColunaPrecisao = 255
            Else
                .lColunaTamanho = tColunasTabelas.lColunaTamanho
                .lColunaPrecisao = tColunasTabelas.lColunaPrecisao
            End If
            .lTamanhoTela = 1100 + (.lColunaTamanho * 10)
        
        End With

        colColunasTabelas.Add objColunasTabelas

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131703

    Loop

    Call Comando_Fechar(lComando)

    ColunasTabelas_Le = SUCESSO

    Exit Function

Erro_ColunasTabelas_Le:

    ColunasTabelas_Le = gErr

    Select Case gErr

        Case 131700
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 131701, 131702, 131703
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SYSCOLUMNS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143990)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function Carrega_ComboArquivo(ByVal objComboBox As ComboBox) As Long

Dim lErro As Long
Dim lComando As Long
Dim sNomeTabela As String
Dim iSeq As Long

On Error GoTo Erro_Carrega_ComboArquivo

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 131704
    
    sNomeTabela = String(STRING_STRING_MAX, 0)

    'Le o sysobjects filtrados por Views e UserTables
    lErro = Comando_Executar(lComando, "SELECT O.Name FROM sysobjects AS O WHERE O.xtype IN (?,?) ORDER By O.Name", _
                                        sNomeTabela, ARQUIVO_TABELA, ARQUIVO_VIEW)
        
    If lErro <> AD_SQL_SUCESSO Then gError 131705

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131706

    iSeq = 0

    Do While lErro <> AD_SQL_SEM_DADOS
    
        iSeq = iSeq + 1

        objComboBox.AddItem sNomeTabela
        objComboBox.ItemData(objComboBox.NewIndex) = iSeq

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131707

    Loop

    Call Comando_Fechar(lComando)

    Carrega_ComboArquivo = SUCESSO

    Exit Function

Erro_Carrega_ComboArquivo:

    Carrega_ComboArquivo = gErr

    Select Case gErr

        Case 131704
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 131705, 131706, 131707
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SYSOBJECTS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143991)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function BrowseArquivo_Le_Todos(ByVal colBrowseArquivo As Collection) As Long
'le todos os campos da tela de browse para o usuario especificado e coloca os resultados na coleção

Dim lComando As Long
Dim lErro As Long
Dim objBrowseArquivo As AdmBrowseArquivo
Dim tBrowseArquivo As typeBrowseArquivo
    
On Error GoTo Erro_BrowseArquivo_Le_Todos

    tBrowseArquivo.sClasse = String(NOME_CLASSE, 0)
    tBrowseArquivo.sNomeTela = String(STRING_NOME_TELA, 0)
    tBrowseArquivo.sProjeto = String(NOME_PROJETO, 0)
    tBrowseArquivo.sNomeArq = String(STRING_NOME_TABELA, 0)
    tBrowseArquivo.sSelecaoSQL = String(STRING_SELECAO_SQL, 0)
    tBrowseArquivo.sClasseBrowser = String(STRING_NOME_CLASSEBROWSER, 0)
    tBrowseArquivo.sTrataParametros = String(STRING_NOME_TRATAPARAMETROS, 0)
    tBrowseArquivo.sTituloBrowser = String(STRING_TITULOBROWSER, 0)
    tBrowseArquivo.sRotinaBotaoEdita = String(STRING_NOME_ROTINABOTAOEDITA, 0)
    tBrowseArquivo.sRotinaBotaoSeleciona = String(STRING_NOME_ROTINABOTAOSELECIONA, 0)
    tBrowseArquivo.sRotinaBotaoConsulta = String(STRING_NOME_ROTINABOTAOCONSULTA, 0)
    tBrowseArquivo.sClasseObjeto = String(NOME_CLASSE, 0)
    tBrowseArquivo.sProjetoObjeto = String(NOME_PROJETO, 0)
    tBrowseArquivo.sNomeTelaConsulta = String(STRING_NOME_TELA, 0)
    tBrowseArquivo.sNomeTelaEdita = String(STRING_NOME_TELA, 0)

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 9252

    lErro = Comando_Executar(lComando, "SELECT NomeTela, Projeto, Classe, NomeArq, SelecaoSQL, ClasseBrowser, TrataParametros, TituloBrowser, RotinaBotaoEdita, RotinaBotaoSeleciona, RotinaBotaoConsulta, BotaoSeleciona, BotaoEdita, BotaoConsulta, ProjetoObjeto, ClasseObjeto, NomeTelaConsulta, NomeTelaEdita, BancoDados FROM BrowseArquivo ORDER BY NomeTela", tBrowseArquivo.sNomeTela, tBrowseArquivo.sProjeto, tBrowseArquivo.sClasse, tBrowseArquivo.sNomeArq, tBrowseArquivo.sSelecaoSQL, tBrowseArquivo.sClasseBrowser, tBrowseArquivo.sTrataParametros, tBrowseArquivo.sTituloBrowser, tBrowseArquivo.sRotinaBotaoEdita, tBrowseArquivo.sRotinaBotaoSeleciona, tBrowseArquivo.sRotinaBotaoConsulta, tBrowseArquivo.iBotaoSeleciona, tBrowseArquivo.iBotaoEdita, tBrowseArquivo.iBotaoConsulta, tBrowseArquivo.sProjetoObjeto, tBrowseArquivo.sClasseObjeto, tBrowseArquivo.sNomeTelaConsulta, tBrowseArquivo.sNomeTelaEdita, tBrowseArquivo.iBancoDados)
    If lErro <> AD_SQL_SUCESSO Then gError 9253
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 9254
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objBrowseArquivo = New AdmBrowseArquivo
    
        objBrowseArquivo.sNomeTela = tBrowseArquivo.sNomeTela
        objBrowseArquivo.sClasse = tBrowseArquivo.sClasse
        objBrowseArquivo.sProjeto = tBrowseArquivo.sProjeto
        objBrowseArquivo.sNomeArq = tBrowseArquivo.sNomeArq
        objBrowseArquivo.sSelecaoSQL = tBrowseArquivo.sSelecaoSQL
        objBrowseArquivo.sClasseBrowser = tBrowseArquivo.sClasseBrowser
        objBrowseArquivo.sTrataParametros = tBrowseArquivo.sTrataParametros
        objBrowseArquivo.sTituloBrowser = tBrowseArquivo.sTituloBrowser
        objBrowseArquivo.sRotinaBotaoEdita = tBrowseArquivo.sRotinaBotaoEdita
        objBrowseArquivo.sRotinaBotaoSeleciona = tBrowseArquivo.sRotinaBotaoSeleciona
        objBrowseArquivo.sRotinaBotaoConsulta = tBrowseArquivo.sRotinaBotaoConsulta
        objBrowseArquivo.iBotaoSeleciona = tBrowseArquivo.iBotaoSeleciona
        objBrowseArquivo.iBotaoEdita = tBrowseArquivo.iBotaoEdita
        objBrowseArquivo.iBotaoConsulta = tBrowseArquivo.iBotaoConsulta
        objBrowseArquivo.sProjetoObjeto = tBrowseArquivo.sProjetoObjeto
        objBrowseArquivo.sClasseObjeto = tBrowseArquivo.sClasseObjeto
        objBrowseArquivo.sNomeTelaConsulta = tBrowseArquivo.sNomeTelaConsulta
        objBrowseArquivo.sNomeTelaEdita = tBrowseArquivo.sNomeTelaEdita
        objBrowseArquivo.iBancoDados = tBrowseArquivo.iBancoDados
        
        colBrowseArquivo.Add objBrowseArquivo
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 9254
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    BrowseArquivo_Le_Todos = SUCESSO
    
    Exit Function
    
Erro_BrowseArquivo_Le_Todos:

    BrowseArquivo_Le_Todos = Err

    Select Case Err
    
        Case 9252
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 9253, 9254
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEARQUIVO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143992)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Campos_Le_Todos2(ByVal sNomeArq As String, ByVal colCampo As Collection) As Long
'le todos os campos e carrega-os na coleção passada como parametro

Dim lComando As Long
Dim lErro As Long
Dim tCampos As typeCampos
Dim objCampo As New AdmCampos
    
On Error GoTo Erro_Campos_Le_Todos2

    tCampos.sDescricao = String(STRING_DESCRICAO_CAMPO, 0)
    tCampos.sFormatacao = String(STRING_FORMATACAO_CAMPO, 0)
    tCampos.sNome = String(STRING_NOME_CAMPO, 0)
    tCampos.sNomeArq = String(STRING_NOME_TABELA, 0)
    tCampos.sTituloEntradaDados = String(STRING_TITULO_ENTRADA_DADOS_CAMPO, 0)
    tCampos.sTituloGrid = String(STRING_TITULO_GRID_CAMPO, 0)
    tCampos.sValDefault = String(STRING_VALOR_DEFAULT_CAMPO, 0)
    tCampos.sValidacao = String(STRING_VALIDACAO_CAMPO, 0)
    
    lComando = 0
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 55984

    lErro = Comando_Executar(lComando, "SELECT NomeArq, Nome, Descricao, Obrigatorio, Imexivel, Ativo, ValDefault, Validacao, Formatacao, Tipo, Tamanho, Precisao, Decimais, TamExibicao, TituloEntradaDados, TituloGrid, Subtipo, Alinhamento FROM Campos WHERE NomeArq = ? ", tCampos.sNomeArq, tCampos.sNome, tCampos.sDescricao, tCampos.iObrigatorio, tCampos.iImexivel, tCampos.iAtivo, tCampos.sValDefault, tCampos.sValidacao, tCampos.sFormatacao, tCampos.iTipo, tCampos.iTamanho, tCampos.iPrecisao, tCampos.iDecimais, tCampos.iTamExibicao, tCampos.sTituloEntradaDados, tCampos.sTituloGrid, tCampos.iSubTipo, tCampos.iAlinhamento, sNomeArq)
    If lErro <> AD_SQL_SUCESSO Then gError 55985
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 55986
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objCampo = New AdmCampos
    
        objCampo.sNome = tCampos.sNome
        objCampo.sNomeArq = tCampos.sNomeArq
        objCampo.iAtivo = tCampos.iAtivo
        objCampo.iDecimais = tCampos.iDecimais
        objCampo.iImexivel = tCampos.iImexivel
        objCampo.iObrigatorio = tCampos.iObrigatorio
        objCampo.iPrecisao = tCampos.iPrecisao
        objCampo.iTamanho = tCampos.iTamanho
        objCampo.iTamExibicao = tCampos.iTamExibicao
        objCampo.iTipo = tCampos.iTipo
        objCampo.sDescricao = tCampos.sDescricao
        objCampo.sFormatacao = tCampos.sFormatacao
        objCampo.sTituloEntradaDados = tCampos.sTituloEntradaDados
        objCampo.sTituloGrid = tCampos.sTituloGrid
        objCampo.sValDefault = tCampos.sValDefault
        objCampo.sValidacao = tCampos.sValidacao
        objCampo.iSubTipo = tCampos.iSubTipo
        objCampo.iAlinhamento = tCampos.iAlinhamento
        
        colCampo.Add objCampo, objCampo.sNomeArq + objCampo.sNome
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 55989
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    Campos_Le_Todos2 = SUCESSO
    
    Exit Function
    
Erro_Campos_Le_Todos2:

    Campos_Le_Todos2 = gErr

    Select Case gErr
    
        Case 55984
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 55985, 55986, 55989
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143993)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function BrowseUsuarioCampo_Le_Todos(ByVal sNomeTela As String, ByVal colBrowseUsuarioCampo As Collection) As Long
'le todos os campos da tela de browse para o usuario especificado e coloca os resultados na coleção

Dim lComando As Long
Dim lErro As Long
Dim tBrowseUsuarioCampo As typeBrowseUsuarioCampo
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
    
On Error GoTo Erro_BrowseUsuarioCampo_Le_Todos

    tBrowseUsuarioCampo.sCodUsuario = String(STRING_USUARIO, 0)
    tBrowseUsuarioCampo.sNome = String(STRING_NOME_CAMPO, 0)
    tBrowseUsuarioCampo.sNomeArq = String(STRING_NOME_TABELA, 0)
    tBrowseUsuarioCampo.sNomeTela = String(STRING_NOME_TELA, 0)
    tBrowseUsuarioCampo.sTitulo = String(STRING_TITULO_CAMPO, 0)

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 9031

    lErro = Comando_Executar(lComando, "SELECT NomeTela, CodUsuario, NomeArq, Nome, PosicaoTela, Titulo, Largura FROM BrowseUsuarioCampo WHERE NomeTela=? ORDER BY PosicaoTela", tBrowseUsuarioCampo.sNomeTela, tBrowseUsuarioCampo.sCodUsuario, tBrowseUsuarioCampo.sNomeArq, tBrowseUsuarioCampo.sNome, tBrowseUsuarioCampo.iPosicaoTela, tBrowseUsuarioCampo.sTitulo, tBrowseUsuarioCampo.lLargura, sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then gError 9032
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 9033
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objBrowseUsuarioCampo = New AdmBrowseUsuarioCampo
    
        objBrowseUsuarioCampo.sNomeTela = tBrowseUsuarioCampo.sNomeTela
        objBrowseUsuarioCampo.sCodUsuario = tBrowseUsuarioCampo.sCodUsuario
        objBrowseUsuarioCampo.sNomeArq = tBrowseUsuarioCampo.sNomeArq
        objBrowseUsuarioCampo.sNome = tBrowseUsuarioCampo.sNome
        objBrowseUsuarioCampo.iPosicaoTela = tBrowseUsuarioCampo.iPosicaoTela
        objBrowseUsuarioCampo.sTitulo = tBrowseUsuarioCampo.sTitulo
        objBrowseUsuarioCampo.lLargura = tBrowseUsuarioCampo.lLargura

        colBrowseUsuarioCampo.Add objBrowseUsuarioCampo
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 9034
        
    Loop
        
    Call Comando_Fechar(lComando)
    
    BrowseUsuarioCampo_Le_Todos = SUCESSO
    
    Exit Function
    
Erro_BrowseUsuarioCampo_Le_Todos:

    BrowseUsuarioCampo_Le_Todos = gErr

    Select Case gErr
    
        Case 9031
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 9032, 9033, 9034
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEUSUARIOCAMPO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143994)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function GrupoBrowseCampo_Le_Todos(ByVal sNomeTela As String, ByVal colGrupoBrowseCampo As Collection) As Long
'le os campos selecionados para o grupo x tela em questão em coloca-os na coleção

Dim lComando As Long
Dim lErro As Long
Dim tGrupoBrowseCampo As typeGrupoBrowseCampo
Dim objGrupoBrowseCampo As AdmGrupoBrowseCampo
    
On Error GoTo Erro_GrupoBrowseCampo_Le_Todos

    tGrupoBrowseCampo.sCodGrupo = String(STRING_GRUPO, 0)
    tGrupoBrowseCampo.sNome = String(STRING_NOME_CAMPO, 0)
    tGrupoBrowseCampo.sNomeArq = String(STRING_NOME_TABELA, 0)
    tGrupoBrowseCampo.sNomeTela = String(STRING_NOME_TELA, 0)

    lComando = 0
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 9077

    lErro = Comando_Executar(lComando, "SELECT NomeArq, Nome, CodGrupo FROM GrupoBrowseCampo WHERE NomeTela=?", tGrupoBrowseCampo.sNomeArq, tGrupoBrowseCampo.sNome, tGrupoBrowseCampo.sCodGrupo, sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then gError 9078
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 9079
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objGrupoBrowseCampo = New AdmGrupoBrowseCampo
    
        objGrupoBrowseCampo.sCodGrupo = tGrupoBrowseCampo.sCodGrupo
        objGrupoBrowseCampo.sNomeTela = sNomeTela
        objGrupoBrowseCampo.sNomeArq = tGrupoBrowseCampo.sNomeArq
        objGrupoBrowseCampo.sNome = tGrupoBrowseCampo.sNome
        
        colGrupoBrowseCampo.Add objGrupoBrowseCampo
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 9080
        
    Loop
        
    Call Comando_Fechar(lComando)
    
    GrupoBrowseCampo_Le_Todos = SUCESSO
    
    Exit Function
    
Erro_GrupoBrowseCampo_Le_Todos:

    GrupoBrowseCampo_Le_Todos = gErr

    Select Case gErr
    
        Case 9077
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 9078, 9079, 9080
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPOBROWSECAMPO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143995)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function
'FIM DAS ROTINAS DE BD
'####################################################################

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_LOCALIZACAO_PRODUTO1
    Set Form_Load_Ocx = Me
    Caption = "Criação de Browse"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BrowseCria"
    
End Function

Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
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
'**** fim do trecho a ser copiado *****

Private Sub BotaoAcertarData_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim sNomeTabela As String
Dim colColunasTabelas As New Collection
Dim objColuna As ClassColunasTabelas

On Error GoTo Erro_BotaoAcertarData_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    Open CurDir & "\AcertaDefaultData.sql" For Output As #2
    
    Print #2, "CREATE DEFAULT [FORPRINT_DATA_NULA] AS {d '1822-09-07'}"
    Print #2, "GO"

    For iIndice = 0 To NomeArq.ListCount - 1
    
        sNomeTabela = NomeArq.List(iIndice)
              
        Set colColunasTabelas = New Collection
              
        lErro = ColunasTabelas_Le(sNomeTabela, colColunasTabelas)
        If lErro <> SUCESSO Then gError 131930
        
        For Each objColuna In colColunasTabelas
        
            If objColuna.sArquivoTipo <> ARQUIVO_TABELA Then Exit For
            
            If UCase(objColuna.sColunaTipo) = "DATETIME" Then
            
                Print #2, ""
                Print #2, "EXEC sp_bindefault N'[dbo].[FORPRINT_DATA_NULA]', N'[" & sNomeTabela & "].[" & objColuna.sColuna & "]' "
                Print #2, "GO"
                Print #2, ""
                Print #2, "UPDATE " & sNomeTabela & " SET " & objColuna.sColuna & " = {d '1822-09-07'} WHERE " & objColuna.sColuna & " = {d '1822-07-09'}"
                Print #2, "GO"
            
            End If
        
        Next
    
    Next

    Close #2

    GL_objMDIForm.MousePointer = vbDefault
    MsgBox "O Arquivo foi exportado para " & CurDir & " com o nome de AcertaDefaultData.sql", vbOKOnly, "SGE"
    
    Exit Sub
    
Erro_BotaoAcertarData_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143996)

    End Select

    Close #2
    
    Exit Sub
    
End Sub

Private Sub objEvento_evSelecao(obj1 As Object)

Dim lErro As Long

On Error GoTo Erro_objEvento_evSelecao

    Me.Show

    Exit Sub

Erro_objEvento_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143997)

    End Select

    Exit Sub

End Sub

Private Sub BotaoIncluirTab_Click()

Dim objTela As New ClassCriaTela
Dim objColunasTabelas As ClassColunasTabelas
Dim colColunasTabelas As New Collection
Dim objControle As ClassCriaControles
Dim objControleAux As ClassCriaControles
Dim iIndex As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoIncluirTab_Click

'    lErro = Critica_Tela()
'    If lErro <> SUCESSO Then gError 131712

    lErro = Move_Tela_Memoria(colColunasTabelas)
    If lErro <> SUCESSO Then gError 131713
    
    For Each objColunasTabelas In colColunasTabelas
    
        'Só cria o controle se ele não for o NumInt ou o FilialEmpresa
        If InStr(1, objColunasTabelas.sColuna, "NumInt") = 0 And InStr(1, objColunasTabelas.sColuna, "FilialEmpresa") = 0 Then
            
            'Para campos com tamanho maior que 50 usa um textbox, senão um MaskEditBox
            If objColunasTabelas.lColunaTamanho <= 50 Then
                            
                'Se o campo for do tipo data => Cria UpDown e MaskEdit Apropriados
                If objColunasTabelas.sColunaTipo = "datetime" Then
                
                    Set objControle = New ClassCriaControles
                    iIndex = iIndex + 1
                    objControle.iOrdem = iIndex
                    objControle.sNome = objColunasTabelas.sColuna
                    objControle.sTipo = "MaskEdBox"
                    objControle.iTipo = TIPO_OUTRO
                    
                    objTela.colControles.Add objControle
                    
                    Set objControle = New ClassCriaControles
                    iIndex = iIndex + 1
                    objControle.iOrdem = iIndex
                    objControle.sNome = "UpDown" & objColunasTabelas.sColuna
                    objControle.sTipo = "UpDown"
                    objControle.iTipo = TIPO_OUTRO
                
                    objTela.colControles.Add objControle
               
                Else 'MaskEditBox
                
                    Set objControle = New ClassCriaControles
                    iIndex = iIndex + 1
                    objControle.iOrdem = iIndex
                    objControle.sNome = objColunasTabelas.sColuna
                    objControle.sTipo = "MaskEdBox"
                    objControle.iTipo = TIPO_OUTRO
                    
                    objTela.colControles.Add objControle
                    
                End If
            
            Else 'TextBox
                
                Set objControle = New ClassCriaControles
                iIndex = iIndex + 1
                objControle.iOrdem = iIndex
                objControle.sNome = objColunasTabelas.sColuna
                objControle.sTipo = "TextBox"
                objControle.iTipo = TIPO_OUTRO
                
                objTela.colControles.Add objControle
            
            End If
            
            Set objControle = New ClassCriaControles
            iIndex = iIndex + 1
            objControle.iOrdem = iIndex
            objControle.sNome = "Label" & objColunasTabelas.sColuna
            objControle.sTipo = "Label"
            objControle.iTipo = TIPO_OUTRO
     
            objTela.colControles.Add objControle
     
        End If
            
    Next
    
    For Each objControle In gobjTela.colControles
    
        If objControle.iTipo <> TIPO_OUTRO Then
            objTela.colControles.Add objControle
        End If
    
        For Each objControleAux In objTela.colControles
        
            If objControleAux.sNome = objControle.sNome Then
                objControleAux.sFrame = objControle.sFrame
                objControleAux.sGrid = objControle.sGrid
                Exit For
            End If
        Next
        
    Next
    
    Call Chama_Tela_Modal("TelaTab", objTela)
    
    Set gobjTela = objTela

    Exit Sub

Erro_BotaoIncluirTab_Click:

    Select Case gErr
    
        Case 131712, 131713

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143998)

    End Select

    Exit Sub
    
End Sub

Private Sub Limpa_ObjControle(ByVal objControle As ClassCriaControles)

Dim iIndice As Integer

On Error GoTo Erro_Limpa_ObjControle

    For iIndice = 1 To 500
    
        objControle.sScript(iIndice) = ""
    
    Next

    Exit Sub

Erro_Limpa_ObjControle:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143999)

    End Select

    Exit Sub
    
End Sub

Private Sub Imprime_ObjControle(ByVal objControle As ClassCriaControles, Optional sIdentacao As String = "")

Dim iIndice As Integer

On Error GoTo Erro_Imprime_ObjControle

    For iIndice = 1 To 500
    
        If objControle.sScript(iIndice) <> "" Then
    
            Print #1, sIdentacao & objControle.sScript(iIndice)
            
        End If
    
    Next

    Exit Sub

Erro_Imprime_ObjControle:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144000)

    End Select

    Exit Sub
    
End Sub
            
Private Function Cria_Scripts_Tab() As Long

Dim lErro As Long

On Error GoTo Erro_Cria_Scripts_Tab

    Print #1, ""
    Print #1, "Private Sub Opcao_BeforeClick(Cancel As Integer)"
    Print #1, "    Call TabStrip_TrataBeforeClick(Cancel, Opcao)"
    Print #1, "End Sub"
    Print #1, ""
    Print #1, "Private Sub Opcao_Click()"
    Print #1, ""
    Print #1, "    'Se frame selecionado não for o atual"
    Print #1, "    If Opcao.SelectedItem.Index <> iFrameAtual Then"
    Print #1, ""
    Print #1, "        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub"
    Print #1, ""
    Print #1, "        'Esconde o frame atual, mostra o novo"
    Print #1, "        FrameOpcao(Opcao.SelectedItem.Index).Visible = True"
    Print #1, "        FrameOpcao(iFrameAtual).Visible = False"
    Print #1, "        'Armazena novo valor de iFrameAtual"
    Print #1, "        iFrameAtual = Opcao.SelectedItem.Index"
    Print #1, ""
    Print #1, "    End If"
    Print #1, ""
    Print #1, "End Sub"
    
    Cria_Scripts_Tab = SUCESSO

    Exit Function

Erro_Cria_Scripts_Tab:

    Cria_Scripts_Tab = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144001)

    End Select
    
    Exit Function
    
End Function

Private Sub BotaoIncluirGrid_Click()

Dim objTela As New ClassCriaTela
Dim objColunasTabelas As ClassColunasTabelas
Dim colColunasTabelas As New Collection
Dim objControle As ClassCriaControles
Dim objControleAux As ClassCriaControles
Dim iIndex As Integer
Dim lErro As Long
Dim colCombo As New Collection
Dim iIndice As Integer

On Error GoTo Erro_BotaoIncluirGrid_Click

'    lErro = Critica_Tela()
'    If lErro <> SUCESSO Then gError 131712

    lErro = Move_Tela_Memoria(colColunasTabelas)
    If lErro <> SUCESSO Then gError 131713
    
    For Each objColunasTabelas In colColunasTabelas
    
        'Só cria o controle se ele não for o NumInt ou o FilialEmpresa
        If InStr(1, objColunasTabelas.sColuna, "NumInt") = 0 And InStr(1, objColunasTabelas.sColuna, "FilialEmpresa") = 0 Then
            
            'Para campos com tamanho maior que 50 usa um textbox, senão um MaskEditBox
            If objColunasTabelas.lColunaTamanho <= 50 Then
                            
                'Se o campo for do tipo data => Cria UpDown e MaskEdit Apropriados
                If objColunasTabelas.sColunaTipo = "datetime" Then
                
                    Set objControle = New ClassCriaControles
                    iIndex = iIndex + 1
                    objControle.iOrdem = iIndex
                    objControle.sNome = objColunasTabelas.sColuna
                    objControle.sTipo = "MaskEdBox"
                    objControle.iTipo = TIPO_OUTRO
                    
                    objTela.colControles.Add objControle
                    
                    Set objControle = New ClassCriaControles
                    iIndex = iIndex + 1
                    objControle.iOrdem = iIndex
                    objControle.sNome = "UpDown" & objColunasTabelas.sColuna
                    objControle.sTipo = "UpDown"
                    objControle.iTipo = TIPO_OUTRO
                
                    objTela.colControles.Add objControle
               
                Else 'MaskEditBox
                
                    Set objControle = New ClassCriaControles
                    iIndex = iIndex + 1
                    objControle.iOrdem = iIndex
                    objControle.sNome = objColunasTabelas.sColuna
                    objControle.sTipo = "MaskEdBox"
                    objControle.iTipo = TIPO_OUTRO
                    
                    objTela.colControles.Add objControle
                    
                End If
            
            Else 'TextBox
                
                Set objControle = New ClassCriaControles
                iIndex = iIndex + 1
                objControle.iOrdem = iIndex
                objControle.sNome = objColunasTabelas.sColuna
                objControle.sTipo = "TextBox"
                objControle.iTipo = TIPO_OUTRO
                
                objTela.colControles.Add objControle
            
            End If
            
            Set objControle = New ClassCriaControles
            iIndex = iIndex + 1
            objControle.iOrdem = iIndex
            objControle.sNome = "Label" & objColunasTabelas.sColuna
            objControle.sTipo = "Label"
            objControle.iTipo = TIPO_OUTRO
     
            objTela.colControles.Add objControle
     
        End If
            
    Next
    
    For Each objControle In gobjTela.colControles
    
        If objControle.iTipo <> TIPO_OUTRO Then
            objTela.colControles.Add objControle
        End If
    
        For Each objControleAux In objTela.colControles
        
            If objControleAux.sNome = objControle.sNome Then
                objControleAux.sFrame = objControle.sFrame
                objControleAux.sGrid = objControle.sGrid
                Exit For
            End If
        Next
        
    Next
    
    For iIndice = 0 To NomeArq.ListCount - 1
        Set objControleAux = New ClassCriaControles
    
        objControleAux.sNome = NomeArq.List(iIndice)
        
        colCombo.Add objControleAux
    Next
    
    Call Chama_Tela_Modal("TelaGrid", objTela, colCombo)
    
    Set gobjTela = objTela

    Exit Sub

Erro_BotaoIncluirGrid_Click:

    Select Case gErr
    
        Case 131712, 131713

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144002)

    End Select

    Exit Sub
    
End Sub

Private Function Cria_Scripts_Grid() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objControle As ClassCriaControles
Dim objControleFilho As ClassCriaControles
Dim sProxErro As String
Dim sPrimeiroErro As String
Dim sUltimoErro As String
Dim colCampos As New Collection

On Error GoTo Erro_Cria_Scripts_Grid

    'INICIALIZA
    For Each objControle In gobjTela.colControles
    
        If objControle.iTipo = TIPO_GRID Then

            Print #1, ""
            Print #1, "Private Function Inicializa_" & objControle.sNome & "(objGrid As AdmGrid) As Long"
            Print #1, ""
            Print #1, "Dim iIndice As Integer"
            Print #1, ""
            Print #1, "    Set objGrid= New AdmGrid"
            Print #1, ""
            Print #1, "    'tela em questão"
            Print #1, "    Set objGrid.objForm = Me"
            Print #1, ""
            Print #1, "    'titulos do grid"
            Print #1, "    objGrid.colColuna.Add (" & """" & """" & ")"


            For Each objControleFilho In objControle.colControles
                Print #1, "    objGrid.colColuna.Add (" & """" & objControleFilho.sNome & """" & ")"
            Next

            Print #1, ""
            Print #1, "    'Controles que participam do Grid"

            For Each objControleFilho In objControle.colControles
                Print #1, "    objGrid.colCampo.Add (" & objControleFilho.sNome & ".Name)"
            Next

            Print #1, ""
            Print #1, "    'Colunas do Grid"

            For Each objControleFilho In objControle.colControles
                Print #1, "    iGrid_" & objControleFilho.sNome & "_Col = " & objControleFilho.iOrdem
            Next

            Print #1, ""
            Print #1, "    objGrid.objGrid = " & objControle.sNome
            Print #1, ""
            Print #1, "    'Todas as linhas do grid"
            Print #1, "    objGrid.objGrid.Rows = <NUM_MAX_ITENS_PODE_TER> + 1"
            Print #1, ""
            Print #1, "    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE"
            Print #1, ""
            Print #1, "    objGrid.iLinhasVisiveis = 8"
            Print #1, ""
            Print #1, "    'Largura da primeira coluna"
            Print #1, "    " & objControle.sNome & ".ColWidth(0) = 400"
            Print #1, ""
            Print #1, "    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL"
            Print #1, ""
            Print #1, "    objGrid.iIncluirHScroll = GRID_INCLUIR_HSCROLL"
            Print #1, ""
            Print #1, "    Call Grid_Inicializa(objGrid)"
            Print #1, ""
            Print #1, "    Inicializa_" & objControle.sNome & " = SUCESSO"
            Print #1, ""
            Print #1, "End Function"
            
        End If
        
    Next

    For Each objControle In gobjTela.colControles
    
        If objControle.iTipo = TIPO_GRID Then
        
            'CÓDIGO DO GRID
            Print #1, ""
            Print #1, "Private Sub " & objControle.sNome & "_Click()"
            Print #1, ""
            Print #1, "Dim iExecutaEntradaCelula As Integer"
            Print #1, ""
            Print #1, "    Call Grid_Click(obj" & objControle.sNome & ", iExecutaEntradaCelula)"
            Print #1, ""
            Print #1, "    If iExecutaEntradaCelula = 1 Then"
            Print #1, "        Call Grid_Entrada_Celula(obj" & objControle.sNome & ", iAlterado)"
            Print #1, "    End If"
            Print #1, ""
            Print #1, "End Sub"
            Print #1, ""
            Print #1, "Private Sub " & objControle.sNome & "_GotFocus()"
            Print #1, "    Call Grid_Recebe_Foco(obj" & objControle.sNome & ")"
            Print #1, "End Sub"
            Print #1, ""
            Print #1, "Private Sub " & objControle.sNome & "_EnterCell()"
            Print #1, "    Call Grid_Entrada_Celula(obj" & objControle.sNome & ", iAlterado)"
            Print #1, "End Sub"
            Print #1, ""
            Print #1, "Private Sub " & objControle.sNome & "_LeaveCell()"
            Print #1, "    Call Saida_Celula(obj" & objControle.sNome & ")"
            Print #1, "End Sub"
            Print #1, ""
            Print #1, "Private Sub " & objControle.sNome & "_KeyPress(KeyAscii As Integer)"
            Print #1, ""
            Print #1, "Dim iExecutaEntradaCelula As Integer"
            Print #1, ""
            Print #1, "    Call Grid_Trata_Tecla(KeyAscii, obj" & objControle.sNome & ", iExecutaEntradaCelula)"
            Print #1, ""
            Print #1, "    If iExecutaEntradaCelula = 1 Then"
            Print #1, "        Call Grid_Entrada_Celula(obj" & objControle.sNome & ", iAlterado)"
            Print #1, "    End If"
            Print #1, ""
            Print #1, "End Sub"
            Print #1, ""
            Print #1, "Private Sub " & objControle.sNome & "_RowColChange()"
            Print #1, "    Call Grid_RowColChange(obj" & objControle.sNome & ")"
            Print #1, "End Sub"
            Print #1, ""
            Print #1, "Private Sub " & objControle.sNome & "_Scroll()"
            Print #1, "    Call Grid_Scroll(obj" & objControle.sNome & ")"
            Print #1, "End Sub"
            Print #1, ""
            Print #1, "Private Sub " & objControle.sNome & "_KeyDown(KeyCode As Integer, Shift As Integer)"
            Print #1, "    Call Grid_Trata_Tecla1(KeyCode, obj" & objControle.sNome & ")"
            Print #1, "End Sub"
            Print #1, ""
            Print #1, "Private Sub " & objControle.sNome & "_LostFocus()"
            Print #1, "    Call Grid_Libera_Foco(obj" & objControle.sNome & ")"
            Print #1, "End Sub"
            
            'CÓDIGO DOS CONTROLES
            For Each objControleFilho In objControle.colControles
            
                Print #1, ""
                Print #1, "Private Sub " & objControleFilho.sNome & "_Change()"
                Print #1, "    iAlterado = REGISTRO_ALTERADO"
                Print #1, "End Sub"
                Print #1, ""
                Print #1, "Private Sub " & objControleFilho.sNome & "_GotFocus()"
                Print #1, "    Call Grid_Campo_Recebe_Foco(obj" & objControle.sNome & ")"
                Print #1, "End Sub"
                Print #1, ""
                Print #1, "Private Sub " & objControleFilho.sNome & "_KeyPress(KeyAscii As Integer)"
                Print #1, "    Call Grid_Trata_Tecla_Campo(KeyAscii, obj" & objControle.sNome & ")"
                Print #1, "End Sub"
                Print #1, ""
                Print #1, "Private Sub " & objControleFilho.sNome & "_Validate(Cancel As Boolean)"
                Print #1, ""
                Print #1, "Dim lErro As Long"
                Print #1, ""
                Print #1, "    Set obj" & objControle.sNome & ".objControle = " & objControleFilho.sNome
                Print #1, "    lErro = Grid_Campo_Libera_Foco(obj" & objControle.sNome & ")"
                Print #1, "    If lErro <> SUCESSO Then Cancel = True"
                Print #1, ""
                Print #1, "End Sub"
            
            Next
            
            'CÓDIGO DOS SAIDA DE CÉLULA DOS CONTROLES
            For Each objControleFilho In objControle.colControles

                Print #1, ""
                Print #1, "Private Function Saida_Celula_" & objControleFilho.sNome & "(objGridInt As AdmGrid) As Long"
                Print #1, "'faz a critica da celula do grid que está deixando de ser a corrente"
                Print #1, ""
                Print #1, "Dim lErro As Long"
                Print #1, ""
                Print #1, "On Error GoTo Erro_Saida_Celula_" & objControleFilho.sNome
                Print #1, ""
                Print #1, "    Set objGridInt.objControle = " & objControleFilho.sNome
                Print #1, ""
                Print #1, "    If (" & objControle.sNome & ".Row - " & objControle.sNome & ".FixedRows) = objGridInt.iLinhasExistentes Then"
                Print #1, "        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1"
                Print #1, "    End If"
                Print #1, ""
                
                Call CalculaProximoErro(sProxErro)
                
                Print #1, "    lErro = Grid_Abandona_Celula(objGridInt)"
                Print #1, "    If lErro <> SUCESSO Then gError " & sProxErro
                Print #1, ""
                Print #1, "    Saida_Celula_" & objControleFilho.sNome & " = SUCESSO"
                Print #1, ""
                Print #1, "    Exit Function"
                Print #1, ""
                Print #1, "Erro_Saida_Celula_" & objControleFilho.sNome & ":"
                Print #1, ""
                Print #1, "    Saida_Celula_" & objControleFilho.sNome & " = gErr"
                Print #1, ""
                Print #1, "    Select Case gErr"
                Print #1, ""
                Print #1, "        Case " & sProxErro
                Print #1, "            Call Grid_Trata_Erro_Saida_Celula(objGridInt)"
                Print #1, ""
                Print #1, "        Case Else"
                
                Call CalculaProximoErro(sProxErro)
                Print #1, "            Call Rotina_Erro(vbOKOnly, " & """" & "ERRO_FORNECIDO_PELO_VB" & """" & ", gErr, Error$, " & sProxErro & ")"
                Print #1, "            Call Grid_Trata_Erro_Saida_Celula(objGridInt)"
                Print #1, ""
                Print #1, "    End Select"
                Print #1, ""
                Print #1, "End Function"

            Next
    
        End If
    
    Next
    
    If gbTemGrid Then
    
        'SAÍDA DE CÉLULA
        Print #1, ""
        Print #1, "Public Function Saida_Celula(objGridInt As AdmGrid) As Long"
        Print #1, "'faz a critica da celula do grid que está deixando de ser a corrente"
        Print #1, ""
        Print #1, "Dim lErro As Long"
        Print #1, "Dim iIndice As Integer"
        Print #1, ""
        Print #1, "On Error GoTo Erro_Saida_Celula"
        Print #1, ""
        Print #1, "    lErro = Grid_Inicializa_Saida_Celula(objGridInt)"
        Print #1, ""
        Print #1, "    If lErro = SUCESSO Then"

        iIndice = 0

        For Each objControle In gobjTela.colControles
        
            If objControle.iTipo = TIPO_GRID Then

                Print #1, ""
                Print #1, "        '" & objControle.sNome
                Print #1, "        If objGridInt.objGrid.Name = " & objControle.sNome & ".Name Then"
                Print #1, "            "
                Print #1, "            'Verifica qual a coluna do Grid em questão"
                Print #1, "            Select Case objGridInt.objGrid.Col"
                Print #1, ""
                
                For Each objControleFilho In objControle.colControles

                    iIndice = iIndice + 1
                    
                    Call CalculaProximoErro(sProxErro)
                    
                    'CADA CONTROLE DO GRID
                    Print #1, ""
                    Print #1, "                Case iGrid_" & objControleFilho.sNome & "_Col"
                    Print #1, ""
                    Print #1, "                    lErro = Saida_Celula_" & objControleFilho.sNome & "(objGridInt)"
                    Print #1, "                    If lErro <> SUCESSO Then gError " & sProxErro
                    
                    If iIndice = 1 Then sPrimeiroErro = sProxErro

                Next

                'CADA GRID
                Print #1, ""
                Print #1, "            End Select"
                Print #1, "                    "
                Print #1, "        End If"

            End If
            
        Next
        
        sUltimoErro = sProxErro

        Call CalculaProximoErro(sProxErro)

        Print #1, ""
        Print #1, "        lErro = Grid_Finaliza_Saida_Celula(objGridInt)"
        Print #1, "        If lErro Then gError " & sProxErro
        Print #1, ""
        Print #1, "    End If"
        Print #1, ""
        Print #1, "    Saida_Celula = SUCESSO"
        Print #1, ""
        Print #1, "    Exit Function"
        Print #1, ""
        Print #1, "Erro_Saida_Celula:"
        Print #1, ""
        Print #1, "    Saida_Celula = gErr"
        Print #1, ""
        Print #1, "    Select Case gErr"
        Print #1, ""
        Print #1, "        Case " & sPrimeiroErro & " To " & sUltimoErro
        Print #1, ""
        Print #1, "        Case " & sProxErro
        Print #1, "            Call Grid_Trata_Erro_Saida_Celula(objGridInt)"
        Print #1, ""
        Print #1, "        Case Else"
        
        Call CalculaProximoErro(sProxErro)
        Print #1, "             Call Rotina_Erro(vbOKOnly, " & """" & "ERRO_FORNECIDO_PELO_VB" & """" & ", gErr, Error$, " & sProxErro & ")"
        Print #1, ""
        Print #1, "    End Select"
        Print #1, ""
        Print #1, "    Exit Function"
        Print #1, ""
        Print #1, "End Function"
    
        'ROTINA GRID ENABLED
        Print #1, ""
        Print #1, "Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)"
        Print #1, ""
        Print #1, "Dim lErro As Long"
        Print #1, ""
        Print #1, "On Error GoTo Erro_Rotina_Grid_Enable"
        Print #1, ""
        Print #1, "    'Pesquisa o controle da coluna em questão"
        Print #1, "    Select Case objControl.Name"
        Print #1, ""
        Print #1, "        Case Else"
        Print #1, "            objControl.Enabled = True"
        Print #1, ""
        Print #1, "    End Select"
        Print #1, ""
        Print #1, "    Exit Sub"
        Print #1, ""
        Print #1, "Erro_Rotina_Grid_Enable:"
        Print #1, ""
        Print #1, "    Select Case gErr"
        Print #1, ""
        Print #1, "        Case Else"
        
        Call CalculaProximoErro(sProxErro)
        Print #1, "            Call Rotina_Erro(vbOKOnly, " & """" & "ERRO_FORNECIDO_PELO_VB" & """" & ", gErr, Error$, " & sProxErro & ")"
        Print #1, ""
        Print #1, "    End Select"
        Print #1, ""
        Print #1, "    Exit Sub"
        Print #1, ""
        Print #1, "End Sub"
        
    End If
    
    For Each objControle In gobjTela.colControles
    
        If objControle.iTipo = TIPO_GRID Then

            lErro = Modelo_Function_Cria("Preenche_" & objControle.sNome & "_Tela", colCampos)
            If lErro <> SUCESSO Then gError 999999

            lErro = Modelo_Function_Cria("Move_" & objControle.sNome & "_Memoria", colCampos)
            If lErro <> SUCESSO Then gError 999999

        End If

    Next
        
    Cria_Scripts_Grid = SUCESSO

    Exit Function

Erro_Cria_Scripts_Grid:

    Cria_Scripts_Grid = gErr

    Select Case gErr
    
        Case 999999

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144006)

    End Select
    
    Exit Function
    
End Function

'Valida Tabela/Filtro
Function Tabela_Le_Generico(ByVal sTabela As String, ByVal sWhere As String) As Long

Dim lComando As Long
Dim lErro As Long
Dim iAux As Integer
Dim sWhereNovo As String
    
On Error GoTo Erro_Tabela_Le_Generico

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 9252
    
    sWhereNovo = Replace(sWhere, "?", "''")

    lErro = Comando_Executar(lComando, "SELECT 1 FROM " & sTabela & " WHERE " & sWhereNovo, iAux)
    If lErro <> AD_SQL_SUCESSO Then
        gError 9253
    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 9254
        
    Call Comando_Fechar(lComando)
    
    Tabela_Le_Generico = SUCESSO
    
    Exit Function
    
Erro_Tabela_Le_Generico:

    Tabela_Le_Generico = gErr

    Select Case gErr
    
        Case 9252
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 9253, 9254
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144007)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function BrowseParamSelecao_Le(ByVal sNomeTela As String, ByVal colBrowseParamSelecao As Collection) As Long
'le os valores dos parametros de selecao relacionados a sNomeTela coloca os resultados na coleção

Dim lComando As Long
Dim lErro As Long
Dim tBrowseParamSelecao As typeBrowseParamSelecao
Dim objBrowseParamSelecao As AdmBrowseParamSelecao
    
On Error GoTo Erro_BrowseParamSelecao_Le

    tBrowseParamSelecao.sNomeTela = String(STRING_NOME_TELA, 0)
    tBrowseParamSelecao.sClasse = String(NOME_CLASSE, 0)
    tBrowseParamSelecao.sProjeto = String(NOME_PROJETO, 0)
    tBrowseParamSelecao.sProperty = String(STRING_NOME_CAMPO, 0)

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 89978

    lErro = Comando_Executar(lComando, "SELECT Projeto, Classe, Property FROM BrowseParamSelecao WHERE NomeTela=? ORDER BY Ordem", tBrowseParamSelecao.sProjeto, tBrowseParamSelecao.sClasse, tBrowseParamSelecao.sProperty, sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then gError 89979
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89980
    
    Do While lErro = SUCESSO
    
        Set objBrowseParamSelecao = New AdmBrowseParamSelecao
    
        objBrowseParamSelecao.iOrdem = tBrowseParamSelecao.iOrdem
        objBrowseParamSelecao.sClasse = tBrowseParamSelecao.sClasse
        objBrowseParamSelecao.sNomeTela = tBrowseParamSelecao.sNomeTela
        objBrowseParamSelecao.sProjeto = tBrowseParamSelecao.sProjeto
        objBrowseParamSelecao.sProperty = tBrowseParamSelecao.sProperty
    
        colBrowseParamSelecao.Add objBrowseParamSelecao
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89981
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    BrowseParamSelecao_Le = SUCESSO
    
    Exit Function
    
Erro_BrowseParamSelecao_Le:

    BrowseParamSelecao_Le = gErr

    Select Case gErr
    
        Case 89978
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89979, 89980, 89981
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEPARAMSELECAO", gErr, sNomeTela)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144008)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Sub BotaoValidarCodigo_Click()

Dim lErro As Long, bCaseErr As Boolean, lLinha As Integer, vErro As Variant, bPula As Boolean
Dim objFSO As FileSystemObject, objFolder As Folder, objFile As File, objTS As TextStream
Dim sTipoArq As String, bArqAberto As Boolean, sRegistro As String, sDiretorio As String
Dim sNomeFunc As String, iRegPos As Integer, sCaracter As String, iPos1 As Integer, iPos2 As Integer
Dim colErrosNum As Collection, sAux As String, sFuncAux As String, sErroFunc As String
Dim lTransacao As Long, iIndice As Integer, alComando(0 To 4) As Long, sAux2 As String
Dim dtData As Date, lSeqData As Long, lSeq As Long, colErrosMsg As Collection, bTO As Boolean
Dim colLinha As Collection, objErro As AdmFiltro, sProjeto As String, sNomeTela As String
Dim objFrmAguarde As New ClassFrmAguarde, sProjetoAux As String
Dim objFrmAguardeTela As New FrmAguarde, iAux As Integer, sPos As String

On Error GoTo Erro_BotaoValidarCodigo_Click

sPos = "01"

    'Abertura de transação
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then gError 196874

sPos = "02"

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 196868
    Next
    
sPos = "03"
    
    dtData = Date
    sDiretorio = "C:\Contab\"
    Set objFSO = New FileSystemObject
    lSeq = 0
    
sPos = "04"
    
    lErro = Comando_Executar(alComando(0), "SELECT MAX(SeqData) + 1 FROM ValidaCodigo WHERE Data = ? ", lSeqData, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 196869
     
sPos = "05"
     
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196870

sPos = "06"

    'Pega todos os aquivos da pasta
    Set objFolder = objFSO.GetFolder(sDiretorio)
    
    objFrmAguarde.iTotalItens = objFolder.Files.Count
    Call objFrmAguardeTela.Inicializa_Progressao(objFrmAguarde)
    
sPos = "07"
    
    For Each objFile In objFolder.Files
        
        sNomeFunc = ""
        bCaseErr = False
        lLinha = 0
        Set colErrosNum = New Collection
        Set colErrosMsg = New Collection
        Set colLinha = New Collection
        sTipoArq = UCase(right(objFile.ShortName, 3))
        sProjeto = ""
        sNomeTela = ""

sPos = "08"
        
        For iAux = 1 To 2
        
sPos = "09:" & CStr(iAux) & " Arq: " & objFile.Name
        
            If sProjeto = "" And iAux = 2 Then
                iPos1 = InStr(1, UCase(objFile.Name), "SELECT")
                iPos2 = InStr(1, UCase(objFile.Name), "GRAVA")
                If iPos1 + iPos2 > 0 Then
                    If InStr(1, UCase(objFile.Name), "CLASSSELECT") > 0 Or InStr(1, UCase(objFile.Name), "CLASSGRAVA") > 0 Then
                        sProjeto = "RotinasContab"
                    ElseIf InStr(1, UCase(objFile.Name), "DICSELECT") > 0 Or InStr(1, UCase(objFile.Name), "DICGRAVA") > 0 Then
                        sProjeto = "DicRotinas"
                    ElseIf iPos1 <= 6 And iPos2 <= 6 Then
                        If iPos1 > 0 Then
                            sAux2 = Mid(objFile.Name, 1, iPos1 - 1)
                        Else
                            sAux2 = Mid(objFile.Name, 1, iPos2 - 1)
                        End If
                        sProjeto = "Rotinas" & sAux2
                    Else
                        If iPos1 > 0 Then
                            sAux2 = Mid(objFile.Name, 6, iPos1 - 6)
                        Else
                            sAux2 = Mid(objFile.Name, 6, iPos2 - 6)
                        End If
                        sProjeto = "Rotinas" & sAux2
                    End If
                End If
            
            End If
        
sPos = "10:" & CStr(iAux) & " Arq: " & objFile.Name
        
            If (sTipoArq = "CLS" Or sTipoArq = "CTL" Or sTipoArq = "BAS" Or sTipoArq = "FRM") And objFile.Name <> "BrowseCriaOcx.ctl" Then
            
                bArqAberto = False
                'abrir arquivo texto
                Set objTS = objFSO.OpenTextFile(objFile.Name, 1, 0)
                bArqAberto = True
                
sPos = "11:" & CStr(iAux) & " Arq: " & objFile.Name
                
                'Até chegar ao fim do arquivo
                Do While Not objTS.AtEndOfStream
                
                    'Busca o próximo registro do arquivo
                    sRegistro = Trim(objTS.ReadLine)
                     
sPos = "12:" & CStr(iAux) & " Arq: " & objFile.Name & " Linha: " & sRegistro
                     
                    'Pega o Primeiro Caracter
                    For iRegPos = 1 To Len(sRegistro)
                        sCaracter = Mid(sRegistro, iRegPos, 1)
                        If sCaracter <> " " Then Exit For
                    Next
                    
                    'Se não é um comentário
                    If sCaracter <> "'" Then
                        
sPos = "13:" & CStr(iAux) & " Arq: " & objFile.Name & " Linha: " & sRegistro
                        
                        iPos1 = InStr(1, UCase(sRegistro), "FUNCTION ")
                        iPos2 = InStr(1, UCase(sRegistro), "SUB ")
                        'Se é uma função ou uma sub
                        If InStr(1, UCase(sRegistro), "DECLARE") = 0 And (iPos2 + iPos1) > 0 And left(UCase(sRegistro), 4) <> "END " Then
                        
                            If iPos1 <> 0 Then
                                sAux = Mid(sRegistro, iPos1 + Len("FUNCTION "))
                            Else
                                sAux = Mid(sRegistro, iPos2 + Len("SUB "))
                            End If
                            iRegPos = InStr(1, sAux, "(") - 1
                            If iRegPos = -1 Then iRegPos = Len(sAux)
                            sFuncAux = left(sAux, iRegPos)
                        
                            If Len(sFuncAux) > 0 Then
                                sNomeFunc = sFuncAux
                                lLinha = 0
                                Set colLinha = New Collection
                            End If
                        End If
                        
sPos = "14:" & CStr(iAux) & " Arq: " & objFile.Name & " Linha: " & sRegistro
                        
                        If UCase(sNomeFunc) = "NAME" And iAux = 1 Then
                        
                            iPos1 = InStr(1, UCase(sRegistro), "NAME = ")
                            If iPos1 > 0 Then
                                sAux = Mid(sRegistro, iPos1 + Len("NAME = ") + 1)
                                iPos1 = InStr(1, UCase(sAux), "'")
                                If iPos1 = 0 Then
                                    sNomeTela = left(sAux, Len(sAux) - 1)
                                Else
                                    sNomeTela = left(sAux, iPos1 - 1)
                                End If
                                
                                sProjetoAux = String(255, 0)
                                
                                lErro = Comando_Executar(alComando(4), "SELECT Projeto_Original FROM Telas WHERE Nome = ? ", sProjetoAux, sNomeTela)
                                If lErro <> AD_SQL_SUCESSO Then gError 196869
                                 
                                lErro = Comando_BuscarPrimeiro(alComando(4))
                                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196870
                                If lErro <> AD_SQL_SUCESSO Then
                                    sProjeto = ""
                                Else
                                    sProjeto = sProjetoAux
                                End If
                                
                                If iAux = 1 Then Exit Do
                                
                            End If
                        End If
                        
sPos = "15:" & CStr(iAux) & " Arq: " & objFile.Name & " Linha: " & sRegistro
                        
                        iPos1 = InStr(1, UCase(sRegistro), "END FUNCTION")
                        iPos2 = InStr(1, UCase(sRegistro), "END SUB")
                        If (iPos2 + iPos1) > 0 Then
                                                    
                            '1 = Só testa e pega algumas informações
                            If iAux <> 1 Then
                            
                                'Tem que gravar os dados obtidos na função
                                For Each objErro In colErrosNum
                                    lSeq = lSeq + 1
                                    lErro = Comando_Executar(alComando(2), "INSERT INTO ValidaCodigo (Data,SeqData,Seq,NomeArq,NomeFunc,Linha,Tipo,Erro,TextoLinha,Projeto,Tela) VALUES (?,?,?,?,?,?,?,?,?,?,?) ", _
                                    dtData, lSeqData, lSeq, objFile.Name, sNomeFunc, objErro.vValor, 1, objErro.sCampo, CStr(colLinha.Item(objErro.vValor)), sProjeto, sNomeTela)
                                    If lErro <> AD_SQL_SUCESSO Then
                                        gError 196869
                                    End If
                                Next
        
                                For Each objErro In colErrosMsg
                                    lSeq = lSeq + 1
                                    lErro = Comando_Executar(alComando(3), "INSERT INTO ValidaCodigo (Data,SeqData,Seq,NomeArq,NomeFunc,Linha,Tipo,Erro,TextoLinha,Projeto,Tela) VALUES (?,?,?,?,?,?,?,?,?,?,?) ", _
                                    dtData, lSeqData, lSeq, objFile.Name, sNomeFunc, objErro.vValor, 2, objErro.sCampo, CStr(colLinha.Item(objErro.vValor)), sProjeto, sNomeTela)
                                    If lErro <> AD_SQL_SUCESSO Then
                                        gError 196869
                                    End If
                                Next
                                
                            End If
                            
                            sNomeFunc = ""
                            bCaseErr = False
                            lLinha = 0
                            Set colErrosNum = New Collection
                            Set colErrosMsg = New Collection
                            Set colLinha = New Collection
                        End If
                        
sPos = "16:" & CStr(iAux) & " Arq: " & objFile.Name & " Linha: " & sRegistro
                        
                        If Len(sNomeFunc) > 0 And iAux <> 1 Then
                        
                            lLinha = lLinha + 1
                            colLinha.Add sRegistro
    
                            'Se está dentro de uma função ou sub
                            
                            iPos1 = InStr(1, UCase(sRegistro), "THEN GERROR ")
                            iPos2 = InStr(1, UCase(sRegistro), "THEN ERROR ")
                            
                            'Está gerando um erro, então tem que guardar para ver se está sendo tratado depois
                            If iPos1 + iPos2 > 0 Then
                                If iPos1 <> 0 Then
                                    sAux = Mid(sRegistro, iPos1 + Len("THEN GERROR "))
                                Else
                                    sAux = Mid(sRegistro, iPos2 + Len("THEN ERROR "))
                                End If
                                iRegPos = InStr(1, sAux, "'") - 1
                                If iRegPos = -1 Then iRegPos = InStr(1, sAux, " ") - 1
                                If iRegPos = -1 Then iRegPos = Len(sAux)
                                sErroFunc = left(sAux, iRegPos)
                                Set objErro = New AdmFiltro
                                objErro.sCampo = Trim(Replace(Replace(sErroFunc, ")", ""), "(", ""))
                                objErro.vValor = lLinha
                                colErrosNum.Add objErro
                            End If
                            
sPos = "17:" & CStr(iAux) & " Arq: " & objFile.Name & " Linha: " & sRegistro
                            
                            If bCaseErr Then
                            
                                iPos1 = InStr(1, UCase(sRegistro), "CASE ")
                                If iPos1 > 0 Then
                                    sAux = UCase(Mid(sRegistro, iPos1 + Len("CASE ")))
                                    sAux2 = ""
                                    bTO = False
                                    'Pega o Primeiro Caracter
                                    For iRegPos = 1 To Len(sAux) + 1
                                        bPula = False
                                        If iRegPos <> Len(sAux) + 1 Then
                                            sCaracter = Mid(sAux, iRegPos, 1)
                                            If sCaracter = "," Then bPula = True
                                            If sCaracter = "'" Then bPula = True
                                            If sCaracter = " " Then bPula = True
                                            If (Len(sAux) > iRegPos + 3) And (iRegPos - 1) > 0 Then
                                                If sCaracter = "T" And Mid(sAux, iRegPos - 1, 4) = " TO " Then bPula = True
                                            End If
                                            If (Len(sAux) > iRegPos + 2) And (iRegPos - 2) > 0 Then
                                                If sCaracter = "O" And Mid(sAux, iRegPos - 2, 4) = " TO " Then bPula = True
                                            End If
                                        Else
                                            bPula = True
                                        End If
                                        If Not bPula Then
                                            sAux2 = sAux2 & sCaracter
                                        Else
                                            If sAux2 <> "" Then
                                                'Remove o erro da coleção de erros não tratados
                                                For iIndice = colErrosNum.Count To 1 Step -1
                                                    vErro = colErrosNum.Item(iIndice).sCampo
                                                    If vErro = sAux2 Then colErrosNum.Remove (iIndice)
                                                Next
                                                If bTO Then
                                                    'Remove o erro da coleção de erros não tratados
                                                    For iIndice = colErrosNum.Count To 1 Step -1
                                                        vErro = colErrosNum.Item(iIndice).sCampo
                                                        If vErro < sAux2 And vErro > sErroFunc Then colErrosNum.Remove (iIndice)
                                                    Next
                                                End If
                                                bTO = False
                                                If Len(sAux) >= iRegPos + 3 Then
                                                    If " TO " = Mid(sAux, iRegPos, 4) Then
                                                        bTO = True
                                                        sErroFunc = sAux2
                                                    End If
                                                End If
                                            End If
                                            sAux2 = ""
                                        End If
                                        If sCaracter = "'" Then Exit For
                                    Next
                                End If
                            
                            End If
                            
sPos = "18:" & CStr(iAux) & " Arq: " & objFile.Name & " Linha: " & sRegistro
                            
                            iPos1 = InStr(1, UCase(sRegistro), "SELECT CASE GERR")
                            iPos2 = InStr(1, UCase(sRegistro), "SELECT CASE ERR")
                            'Verifca se está no tratamento de erros
                            If iPos1 + iPos2 > 0 Then
                                bCaseErr = True
                            End If
                            
                            If bCaseErr Then
                                iPos1 = InStr(1, UCase(sRegistro), "ROTINA_ERRO")
                                iPos2 = InStr(1, UCase(sRegistro), "ROTINA_AVISO")
                                'Verifica chamando a mensagem de erro
                                If (iPos1 + iPos2) > 0 And InStr(iPos1 + iPos2 + 1, sRegistro, """") > 0 Then
                                    'Tem que obter o segundo parametro para ver se está cadastrado no BD
                                    sAux = Mid(sRegistro, InStr(iPos1 + iPos2 + 1, sRegistro, """") + 1)
                                    sAux2 = left(sAux, InStr(1, sAux, """") - 1)
                                    
                                    sErroFunc = String(255, 0)
                                    lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM Erros WHERE Codigo = ? ", sErroFunc, sAux2)
                                    If lErro <> AD_SQL_SUCESSO Then gError 196869
                                     
                                    lErro = Comando_BuscarPrimeiro(alComando(1))
                                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196870
                                    
                                    If lErro <> AD_SQL_SUCESSO Then
                                        Set objErro = New AdmFiltro
                                        objErro.sCampo = Trim(sAux2)
                                        objErro.vValor = lLinha
                                        colErrosMsg.Add objErro
                                    End If
                                End If
                            
                            End If
                        
sPos = "20:" & CStr(iAux) & " Arq: " & objFile.Name & " Linha: " & sRegistro
                        
                        End If
                        
                    End If
                     
                Loop
                
                'fechar arquivo texto
                objTS.Close
                bArqAberto = False
            
            End If
        
        Next
        
sPos = "21" & " Arq: " & objFile.Name
        
        Call objFrmAguardeTela.ProcessouItem
        If objFrmAguarde.iCancelar = MARCADO Then Exit For
        
    Next
    
    Set objFrmAguardeTela = Nothing
    If objFrmAguarde.iCancelar = MARCADO Then gError 192084
    
    'Confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then gError 196883
       
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call MsgBox("Rotina executada com sucesso. Resultado na tabela ValidaCodigo no Dic. 1 = Erro não tratado e 2 = Erro não cadastrado no BD", vbOKOnly)
     
    Exit Sub
    
Erro_BotaoValidarCodigo_Click:
     
    Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130555)
    
    Call Rotina_Erro(vbOKOnly, "Posição: " & sPos, gErr)
    
    If bArqAberto Then objTS.Close
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_RollbackDic
    
    If Not (objFrmAguardeTela Is Nothing) Then
        Call objFrmAguardeTela.Trata_Erro
    End If
     
    Exit Sub
    
End Sub
