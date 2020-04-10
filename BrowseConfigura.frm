VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form BrowseConfigura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configura "
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "BrowseConfigura.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3330
      Index           =   5
      Left            =   135
      TabIndex        =   45
      Top             =   495
      Visible         =   0   'False
      Width           =   7245
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2880
         Index           =   0
         Left            =   105
         TabIndex        =   48
         Top             =   360
         Width           =   6975
         Begin VB.Frame Frame4 
            Caption         =   "Formato"
            Height          =   825
            Left            =   135
            TabIndex        =   71
            Top             =   300
            Width           =   6570
            Begin VB.CheckBox NomeAuto 
               Caption         =   "Nome Automático"
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
               Left            =   330
               TabIndex        =   54
               Top             =   525
               Value           =   1  'Checked
               Width           =   1830
            End
            Begin VB.TextBox Arquivo 
               Enabled         =   0   'False
               Height          =   285
               Left            =   3465
               MaxLength       =   255
               TabIndex        =   55
               Top             =   480
               Width           =   2625
            End
            Begin VB.CommandButton BotaoProcurar 
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
               Left            =   6105
               TabIndex        =   56
               Top             =   150
               Width           =   375
            End
            Begin VB.TextBox LocalizacaoCSV 
               Height          =   285
               Left            =   3465
               MaxLength       =   255
               TabIndex        =   53
               Top             =   150
               Width           =   2625
            End
            Begin VB.OptionButton FormatoCSV 
               Caption         =   "CSV"
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
               Left            =   1350
               TabIndex        =   52
               Top             =   210
               Width           =   795
            End
            Begin VB.OptionButton FormatoXls 
               Caption         =   "XLS"
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
               Left            =   330
               TabIndex        =   51
               Top             =   210
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.Label Label12 
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
               Height          =   225
               Left            =   2640
               TabIndex        =   73
               Top             =   525
               Width           =   720
            End
            Begin VB.Label Label11 
               Caption         =   "Localização:"
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
               Left            =   2280
               TabIndex        =   72
               Top             =   195
               Width           =   1710
            End
         End
         Begin VB.Frame FrameForm 
            Caption         =   "Fórmulas"
            Height          =   1770
            Left            =   135
            TabIndex        =   57
            Top             =   1110
            Width           =   6570
            Begin VB.ComboBox FormFormulas 
               Height          =   315
               ItemData        =   "BrowseConfigura.frx":014A
               Left            =   3450
               List            =   "BrowseConfigura.frx":015D
               Style           =   2  'Dropdown List
               TabIndex        =   60
               Top             =   840
               Width           =   2145
            End
            Begin VB.ComboBox FormCampos 
               Height          =   315
               Left            =   1290
               Style           =   2  'Dropdown List
               TabIndex        =   59
               Top             =   870
               Width           =   3135
            End
            Begin MSFlexGridLib.MSFlexGrid GridFormulas 
               Height          =   1365
               Left            =   270
               TabIndex        =   58
               Top             =   195
               Width           =   6045
               _ExtentX        =   10663
               _ExtentY        =   2408
               _Version        =   393216
               Rows            =   10
               Cols            =   4
               BackColorSel    =   -2147483643
               ForeColorSel    =   -2147483640
               AllowBigSelection=   0   'False
               FocusRect       =   2
            End
         End
         Begin VB.TextBox PlanTitulo 
            Height          =   285
            Left            =   915
            MaxLength       =   50
            TabIndex        =   50
            Top             =   15
            Width           =   5730
         End
         Begin VB.Label Label8 
            Caption         =   "Título:"
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
            Left            =   180
            TabIndex        =   49
            Top             =   45
            Width           =   1035
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2865
         Index           =   1
         Left            =   60
         TabIndex        =   47
         Top             =   360
         Visible         =   0   'False
         Width           =   7050
         Begin VB.CheckBox IncluiTabela 
            Caption         =   "Incluir tabela dinâmica"
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
            Left            =   60
            TabIndex        =   70
            Top             =   300
            Width           =   2340
         End
         Begin VB.Frame FrameGraf 
            Caption         =   "Gráfico"
            Enabled         =   0   'False
            Height          =   555
            Left            =   2430
            TabIndex        =   66
            Top             =   0
            Width           =   4590
            Begin VB.CheckBox IncluirGrafico 
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
               Height          =   195
               Left            =   75
               TabIndex        =   68
               Top             =   255
               Width           =   945
            End
            Begin VB.ComboBox TipoGrafico 
               Height          =   315
               ItemData        =   "BrowseConfigura.frx":0187
               Left            =   1605
               List            =   "BrowseConfigura.frx":0194
               Style           =   2  'Dropdown List
               TabIndex        =   67
               Top             =   195
               Width           =   2925
            End
            Begin VB.Label Label10 
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
               Height          =   225
               Left            =   1080
               TabIndex        =   69
               Top             =   255
               Width           =   1035
            End
         End
         Begin VB.Frame FrameCampos 
            Caption         =   "Campos"
            Height          =   2040
            Left            =   60
            TabIndex        =   61
            Top             =   660
            Width           =   6945
            Begin VB.ComboBox TabCamposPosicao 
               Height          =   315
               ItemData        =   "BrowseConfigura.frx":01AE
               Left            =   1920
               List            =   "BrowseConfigura.frx":01C1
               Style           =   2  'Dropdown List
               TabIndex        =   65
               Top             =   1245
               Width           =   1395
            End
            Begin VB.ComboBox TabCamposForm 
               Height          =   315
               ItemData        =   "BrowseConfigura.frx":01EB
               Left            =   660
               List            =   "BrowseConfigura.frx":01FE
               Style           =   2  'Dropdown List
               TabIndex        =   64
               Top             =   690
               Width           =   1395
            End
            Begin VB.ComboBox TabCamposCampos 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2775
               Style           =   2  'Dropdown List
               TabIndex        =   63
               Top             =   690
               Width           =   2820
            End
            Begin MSFlexGridLib.MSFlexGrid GridTabCampos 
               Height          =   1695
               Left            =   300
               TabIndex        =   62
               Top             =   225
               Width           =   6360
               _ExtentX        =   11218
               _ExtentY        =   2990
               _Version        =   393216
               Rows            =   10
               Cols            =   4
               BackColorSel    =   -2147483643
               ForeColorSel    =   -2147483640
               AllowBigSelection=   0   'False
               FocusRect       =   2
            End
         End
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   3315
         Left            =   30
         TabIndex        =   46
         Top             =   0
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   5847
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dados"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Tabela Dinâmica"
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   0
      Left            =   135
      TabIndex        =   1
      Top             =   780
      Width           =   7185
      Begin VB.CommandButton BotaoSelecionarTodos 
         Caption         =   "Marcar Todos"
         Height          =   525
         Left            =   4095
         Picture         =   "BrowseConfigura.frx":0228
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   885
         Width           =   1470
      End
      Begin VB.CommandButton BotaoSelecionar 
         Caption         =   "Marcar"
         Height          =   525
         Left            =   4095
         Picture         =   "BrowseConfigura.frx":1242
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   285
         Width           =   1470
      End
      Begin VB.CommandButton BotaoDesmarcar 
         Caption         =   "Desmarcar"
         Height          =   525
         Left            =   4095
         Picture         =   "BrowseConfigura.frx":196C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1470
         Width           =   1470
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   525
         Left            =   4095
         Picture         =   "BrowseConfigura.frx":206E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2070
         Width           =   1470
      End
      Begin VB.ListBox CamposDisponiveis 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   1185
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   495
         Width           =   2400
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Campos Disponíveis"
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
         Left            =   1185
         TabIndex        =   22
         Top             =   210
         Width           =   1740
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Index           =   1
      Left            =   135
      TabIndex        =   7
      Top             =   780
      Visible         =   0   'False
      Width           =   7185
      Begin VB.CommandButton BotaoParaBaixo 
         Height          =   390
         Left            =   4485
         Picture         =   "BrowseConfigura.frx":3250
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1620
         Width           =   390
      End
      Begin VB.CommandButton BotaoParaCima 
         Height          =   390
         Left            =   4470
         Picture         =   "BrowseConfigura.frx":3412
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   975
         Width           =   390
      End
      Begin VB.ListBox CamposPosicionados 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   1530
         TabIndex        =   8
         Top             =   570
         Width           =   2670
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ordem de Exibição dos Campos"
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
         Left            =   1515
         TabIndex        =   23
         Top             =   285
         Width           =   2685
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   3
      Left            =   135
      TabIndex        =   15
      Top             =   780
      Visible         =   0   'False
      Width           =   7185
      Begin VB.ComboBox ComboParFechar 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "BrowseConfigura.frx":35D4
         Left            =   1470
         List            =   "BrowseConfigura.frx":35EA
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   930
         Width           =   585
      End
      Begin VB.ComboBox ComboParAbrir 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "BrowseConfigura.frx":360A
         Left            =   480
         List            =   "BrowseConfigura.frx":3620
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   1005
         Width           =   585
      End
      Begin VB.CommandButton BotaoLimparPesquisa 
         Caption         =   "Limpar Pesquisa"
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
         Left            =   2160
         TabIndex        =   17
         Top             =   2445
         Width           =   2055
      End
      Begin VB.ComboBox ComboConjuncao 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "BrowseConfigura.frx":3640
         Left            =   5460
         List            =   "BrowseConfigura.frx":364A
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1470
         Width           =   615
      End
      Begin VB.TextBox Valor 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   3480
         MaxLength       =   150
         TabIndex        =   26
         Top             =   1500
         Width           =   1815
      End
      Begin VB.ComboBox ComboCampo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "BrowseConfigura.frx":3655
         Left            =   660
         List            =   "BrowseConfigura.frx":3657
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1485
         Width           =   1800
      End
      Begin VB.ComboBox ComboOperacao 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "BrowseConfigura.frx":3659
         Left            =   2565
         List            =   "BrowseConfigura.frx":3672
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1485
         Width           =   900
      End
      Begin MSFlexGridLib.MSFlexGrid GridSelecao 
         Height          =   2025
         Left            =   -30
         TabIndex        =   16
         Top             =   285
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   3572
         _Version        =   393216
         Rows            =   10
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Condição de Seleção"
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
         TabIndex        =   27
         Top             =   60
         Width           =   1830
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2895
      Index           =   4
      Left            =   135
      TabIndex        =   29
      Top             =   780
      Visible         =   0   'False
      Width           =   7185
      Begin VB.CommandButton BotaoLimparOrdenacao 
         Height          =   375
         Left            =   6075
         Picture         =   "BrowseConfigura.frx":3691
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   240
         Width           =   390
      End
      Begin VB.CommandButton BotaoInserirCampoOrdenacao 
         Height          =   330
         Left            =   3600
         Picture         =   "BrowseConfigura.frx":3BC3
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1350
         Width           =   360
      End
      Begin VB.CommandButton BotaoRemoverCampoOrdenacao 
         Height          =   330
         Left            =   3600
         Picture         =   "BrowseConfigura.frx":3D6D
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1965
         Width           =   360
      End
      Begin VB.CommandButton BotaoParaCimaOrdenacao 
         Height          =   330
         Left            =   540
         Picture         =   "BrowseConfigura.frx":3F17
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1350
         Width           =   360
      End
      Begin VB.CommandButton BotaoParaBaixoOrdenacao 
         Height          =   330
         Left            =   540
         Picture         =   "BrowseConfigura.frx":40D9
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1965
         Width           =   360
      End
      Begin VB.CommandButton BotaoInserirOrdenacao 
         Height          =   375
         Left            =   5115
         Picture         =   "BrowseConfigura.frx":429B
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   240
         Width           =   390
      End
      Begin VB.CommandButton BotaoRemoverOrdenacao 
         Height          =   375
         Left            =   5595
         Picture         =   "BrowseConfigura.frx":43F5
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   240
         Width           =   390
      End
      Begin VB.ListBox ListaOrdenacao 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   1050
         TabIndex        =   34
         Top             =   930
         Width           =   2400
      End
      Begin VB.ListBox ListaCamposDisponiveis 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   4080
         Sorted          =   -1  'True
         TabIndex        =   32
         Top             =   930
         Width           =   2400
      End
      Begin VB.ComboBox Ordenacao 
         Height          =   315
         Left            =   525
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   255
         Width           =   4470
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ordenação"
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
         Left            =   1080
         TabIndex        =   35
         Top             =   690
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Campos Disponíveis"
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
         Left            =   4080
         TabIndex        =   33
         Top             =   690
         Width           =   1740
      End
      Begin VB.Label Label5 
         Caption         =   "Ordenações criadas pelo usuário"
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
         Left            =   540
         TabIndex        =   31
         Top             =   45
         Width           =   3210
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   2
      Left            =   135
      TabIndex        =   11
      Top             =   780
      Visible         =   0   'False
      Width           =   7185
      Begin VB.ListBox TitulosCampos 
         Height          =   450
         Left            =   3570
         TabIndex        =   14
         Top             =   1875
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.TextBox Titulo 
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
         Left            =   3420
         MaxLength       =   50
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1125
         Width           =   2580
      End
      Begin VB.ListBox CamposExibidos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Left            =   750
         TabIndex        =   12
         Top             =   405
         Width           =   2295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Título do Campo"
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
         Left            =   3405
         TabIndex        =   21
         Top             =   855
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Campos Exibidos"
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
         Left            =   810
         TabIndex        =   20
         Top             =   120
         Width           =   1440
      End
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      Height          =   495
      Left            =   3765
      Picture         =   "BrowseConfigura.frx":457F
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3945
      Width           =   1230
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1650
      Picture         =   "BrowseConfigura.frx":4681
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3945
      Width           =   1230
   End
   Begin MSComctlLib.TabStrip Opcoes 
      Height          =   3750
      Left            =   15
      TabIndex        =   0
      Top             =   150
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   6615
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Campos Exibidos"
            Object.ToolTipText     =   "Campos Selecionados para Exibição"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Posição dos Campos"
            Object.ToolTipText     =   "Posição de Exibição dos Campos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Títulos"
            Object.ToolTipText     =   "Títulos dos Campos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pesquisa"
            Object.ToolTipText     =   "Seleção das Informações a Serem Exibidas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ordenação"
            Object.ToolTipText     =   "Definição da ordem em que os registros serão exibidos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Excel"
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
Attribute VB_Name = "BrowseConfigura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" _
                                  (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                  (ByVal pidList As Long, _
                                  ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                  (ByVal lpString1 As String, ByVal _
                                  lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

Private colBrowseUsuarioCampo As Collection
Private colCamposDisponiveis As Collection
Private iFrameAtual As Integer
Private iFrame2Atual As Integer
Private objBrowseConfigura1 As AdmBrowseConfigura
Dim objGrid1 As AdmGrid
Dim iAlterado As Integer
Dim iAlteradoOrdenacao As Integer

Const COL_GRIDSELECAO_PAR_ABRIR = 1
Const COL_GRIDSELECAO_CAMPO = 2
Const COL_GRIDSELECAO_OP = 3
Const COL_GRIDSELECAO_VALOR = 4
Const COL_GRIDSELECAO_PAR_FECHAR = 5
Const COL_GRIDSELECAO_EOU = 6

Dim objGridForm As AdmGrid
Const COL_GRIDFORMULAS_CAMPO = 1
Const COL_GRIDFORMULAS_FORMULA = 2

Dim objGridCampos As AdmGrid
Const COL_GRIDCAMPOS_CAMPO = 1
Const COL_GRIDCAMPOS_POSICAO = 2
Const COL_GRIDCAMPOS_FORMULA = 3

Const OP_LIKE As String = "LIKE"
Const TAB_INDEX_PESQUISA = 4
Const TAB_INDEX_CAMPOS_EXIBIDOS = 1

Private Sub BotaoCancela_Click()
    
    objBrowseConfigura1.iTelaOK = CANCELA
    
    Unload BrowseConfigura

End Sub

Private Sub BotaoDesmarcar_Click()

Dim iIndice As Integer

    If CamposDisponiveis.ListIndex > -1 Then
        If CamposDisponiveis.Selected(CamposDisponiveis.ListIndex) = True Then
            CamposDisponiveis.Selected(CamposDisponiveis.ListIndex) = False
        End If
    End If
    
End Sub

Private Sub BotaoDesmarcarTodos_Click()

Dim iIndice As Integer

    For iIndice = 0 To CamposDisponiveis.ListCount - 1
        CamposDisponiveis.Selected(iIndice) = False
    Next
    
End Sub

Private Sub BotaoInserirCampoOrdenacao_Click()

    If ListaCamposDisponiveis.ListIndex <> -1 Then
    
        ListaOrdenacao.AddItem ListaCamposDisponiveis.Text
        ListaCamposDisponiveis.RemoveItem (ListaCamposDisponiveis.ListIndex)
    
    End If
    
End Sub

Private Sub BotaoInserirOrdenacao_Click()

Dim objBrowseIndice As AdmBrowseIndice
Dim objBrowseIndice1 As AdmBrowseIndice
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoInserirOrdenacao_Click

    If ListaOrdenacao.ListCount = 0 Then gError 60830

    iIndice = 0

    'descobre o proximo indice
    For Each objBrowseIndice In objBrowseConfigura1.colBrowseIndiceUsuario
        If objBrowseIndice.iIndice > iIndice Then iIndice = objBrowseIndice.iIndice
    Next
                
    Set objBrowseIndice = New AdmBrowseIndice
                
    objBrowseIndice.iIndice = iIndice + 1
    objBrowseIndice.sNomeTela = objBrowseConfigura1.sNomeTela
    objBrowseIndice.sNomeIndice = ""
    objBrowseIndice.sOrdenacaoSQL = ""
    objBrowseIndice.sSelecaoSQL = ""
    
    For iIndice = 0 To ListaOrdenacao.ListCount - 1
        objBrowseIndice.sNomeIndice = objBrowseIndice.sNomeIndice & ListaOrdenacao.List(iIndice) & " + "
        objBrowseIndice.sOrdenacaoSQL = objBrowseIndice.sOrdenacaoSQL & ListaOrdenacao.List(iIndice) & ","
        objBrowseIndice.sSelecaoSQL = objBrowseIndice.sSelecaoSQL & "("
        For iIndice1 = 0 To iIndice - 1
            objBrowseIndice.sSelecaoSQL = objBrowseIndice.sSelecaoSQL & ListaOrdenacao.List(iIndice1) & "=? AND "
        Next
        objBrowseIndice.sSelecaoSQL = objBrowseIndice.sSelecaoSQL & ListaOrdenacao.List(iIndice) & "<?) OR "
    Next
    
    objBrowseIndice.sNomeIndice = left(objBrowseIndice.sNomeIndice, Len(objBrowseIndice.sNomeIndice) - 3)
    objBrowseIndice.sOrdenacaoSQL = left(objBrowseIndice.sOrdenacaoSQL, Len(objBrowseIndice.sOrdenacaoSQL) - 1)
    objBrowseIndice.sSelecaoSQL = left(objBrowseIndice.sSelecaoSQL, Len(objBrowseIndice.sSelecaoSQL) - 4)

    If Len(objBrowseIndice.sSelecaoSQL) > STRING_ORDENACAO_SQL Then gError 187532

    'pesquisa se o indice já foi criado pelo usuario
    For Each objBrowseIndice1 In objBrowseConfigura1.colBrowseIndiceUsuario
        If objBrowseIndice1.sOrdenacaoSQL = objBrowseIndice.sOrdenacaoSQL Then gError 60848
    Next
    
    'pesquisa se o indice já foi criado pelo sistema
    For Each objBrowseIndice1 In objBrowseConfigura1.colBrowseIndice
        If objBrowseIndice1.sOrdenacaoSQL = objBrowseIndice.sOrdenacaoSQL Then gError 60849
    Next

    objBrowseConfigura1.colBrowseIndiceUsuario.Add objBrowseIndice
    Ordenacao.AddItem objBrowseIndice.sNomeIndice
    Ordenacao.ItemData(Ordenacao.NewIndex) = objBrowseIndice.iIndice
    
    Ordenacao.ListIndex = -1
    
    Call BotaoLimparOrdenacao_Click
    
    objBrowseConfigura1.iAlteradoOrdenacao = 1
    
    Exit Sub
    
Erro_BotaoInserirOrdenacao_Click:

    Select Case gErr
    
        Case 60830
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LISTA_ORDENACAO_VAZIA", gErr)
    
        Case 60848
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDENACAO_JA_CADASTRADA", gErr)
    
        Case 60849
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDENACAO_SISTEMA_JA_CADASTRADA", gErr)
    
        Case 187532
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDENACAO_MUITO_GRANDE", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143884)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoLimparOrdenacao_Click()

Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo

    Ordenacao.ListIndex = -1

    ListaOrdenacao.Clear
    ListaCamposDisponiveis.Clear
    
    'carrega a combo de campos da lista de campos disponiveis para ordenacao
    For Each objBrowseUsuarioCampo In colCamposDisponiveis
        ListaCamposDisponiveis.AddItem objBrowseUsuarioCampo.sNome
    Next
    

End Sub

Private Sub BotaoLimparPesquisa_Click()

    Call Grid_Limpa(objGrid1)

End Sub

Private Sub BotaoOK_Click()
    
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
Dim iIndice As Integer
Dim iIndiceColecao As Integer
Dim iAchei As Integer
Dim lErro As Long
Dim objBrowseExcelAux As AdmBrowseExcelAux

On Error GoTo Erro_BotaoOK_Click

    iIndiceColecao = 0

    For Each objBrowseUsuarioCampo In colBrowseUsuarioCampo
    
        iIndiceColecao = iIndiceColecao + 1
        For iIndice = 0 To CamposPosicionados.ListCount - 1
            If CamposPosicionados.List(iIndice) = objBrowseUsuarioCampo.sNome Then
                objBrowseUsuarioCampo.iPosicaoTela = iIndice + 1
                objBrowseUsuarioCampo.sTitulo = TitulosCampos.List(iIndice)
                Exit For
            End If
        Next
        
        'se não encontrou o campo na lista ==> o campo deve ser removido
        If iIndice = CamposPosicionados.ListCount Then
            colBrowseUsuarioCampo.Remove (iIndiceColecao)
            iIndiceColecao = iIndiceColecao - 1
        End If
    Next
    
    For iIndice = 0 To CamposPosicionados.ListCount - 1
        
        iAchei = 0
        
        For Each objBrowseUsuarioCampo In colBrowseUsuarioCampo
            If CamposPosicionados.List(iIndice) = objBrowseUsuarioCampo.sNome Then
                iAchei = 1
                Exit For
            End If
        Next
        
        If iAchei = 0 Then
            For Each objBrowseUsuarioCampo In colCamposDisponiveis
                If CamposPosicionados.List(iIndice) = objBrowseUsuarioCampo.sNome Then
                    colBrowseUsuarioCampo.Add objBrowseUsuarioCampo
                    objBrowseUsuarioCampo.iPosicaoTela = iIndice + 1
                    objBrowseUsuarioCampo.sTitulo = TitulosCampos.List(iIndice)
                    Exit For
                End If
            Next
        End If
    Next
    
    If colBrowseUsuarioCampo.Count = 0 Then gError 9123
    
    lErro = Move_GridSelecao_Memoria()
    If lErro <> SUCESSO Then gError 20657
    
    objBrowseConfigura1.objBrowseExcel.sTitulo = PlanTitulo.Text
    
    If IncluiTabela.Value = vbChecked Then
        objBrowseConfigura1.objBrowseExcel.iTabelaDinamica = MARCADO
    Else
        objBrowseConfigura1.objBrowseExcel.iTabelaDinamica = DESMARCADO
    End If
    
    If IncluirGrafico.Value = vbChecked Then
        objBrowseConfigura1.objBrowseExcel.iIncluirGrafico = MARCADO
    Else
        objBrowseConfigura1.objBrowseExcel.iIncluirGrafico = DESMARCADO
    End If
    
    If FormatoXls.Value Then
        objBrowseConfigura1.objBrowseExcel.iFormato = EXCEL_FORMATO_XLS
    Else
        objBrowseConfigura1.objBrowseExcel.iFormato = EXCEL_FORMATO_CSV
        If Len(Trim(LocalizacaoCSV.Text)) = 0 Then gError 202232
    End If
    objBrowseConfigura1.objBrowseExcel.sLocalizacaoCsv = LocalizacaoCSV.Text
    
    
    Select Case TipoGrafico.Text
        Case ""
            objBrowseConfigura1.objBrowseExcel.iTipoGrafico = 0
        Case EXCEL_TIPOGRAFICO_AREA_TEXTO
            objBrowseConfigura1.objBrowseExcel.iTipoGrafico = EXCEL_TIPOGRAFICO_AREA
        Case EXCEL_TIPOGRAFICO_LINHA_TEXTO
            objBrowseConfigura1.objBrowseExcel.iTipoGrafico = EXCEL_TIPOGRAFICO_LINHA
        Case EXCEL_TIPOGRAFICO_COLUNA_TEXTO
            objBrowseConfigura1.objBrowseExcel.iTipoGrafico = EXCEL_TIPOGRAFICO_COLUNA
        Case EXCEL_TIPOGRAFICO_PIZZA_TEXTO
            objBrowseConfigura1.objBrowseExcel.iTipoGrafico = EXCEL_TIPOGRAFICO_PIZZA
    End Select
    
    For iIndice = objBrowseConfigura1.objBrowseExcel.colFormulas.Count To 1 Step -1
        objBrowseConfigura1.objBrowseExcel.colFormulas.Remove iIndice
    Next
    
    For iIndice = 1 To objGridForm.iLinhasExistentes
        Set objBrowseExcelAux = New AdmBrowseExcelAux
        objBrowseExcelAux.sCampo = GridFormulas.TextMatrix(iIndice, COL_GRIDFORMULAS_CAMPO)
        Select Case GridFormulas.TextMatrix(iIndice, COL_GRIDFORMULAS_FORMULA)
            Case EXCEL_FORMULA_SUM_TEXTO
                objBrowseExcelAux.iFormula = EXCEL_FORMULA_SUM
            Case EXCEL_FORMULA_COUNT_TEXTO
                objBrowseExcelAux.iFormula = EXCEL_FORMULA_COUNT
            Case EXCEL_FORMULA_MIN_TEXTO
                objBrowseExcelAux.iFormula = EXCEL_FORMULA_MIN
            Case EXCEL_FORMULA_MAX_TEXTO
                objBrowseExcelAux.iFormula = EXCEL_FORMULA_MAX
            Case EXCEL_FORMULA_AVG_TEXTO
                objBrowseExcelAux.iFormula = EXCEL_FORMULA_AVG
        End Select
        objBrowseConfigura1.objBrowseExcel.colFormulas.Add objBrowseExcelAux
    Next
    
    For iIndice = objBrowseConfigura1.objBrowseExcel.colCampos.Count To 1 Step -1
        objBrowseConfigura1.objBrowseExcel.colCampos.Remove iIndice
    Next
    
    For iIndice = 1 To objGridCampos.iLinhasExistentes
        Set objBrowseExcelAux = New AdmBrowseExcelAux
        objBrowseExcelAux.sCampo = GridTabCampos.TextMatrix(iIndice, COL_GRIDCAMPOS_CAMPO)
        Select Case GridTabCampos.TextMatrix(iIndice, COL_GRIDCAMPOS_FORMULA)
            Case ""
                objBrowseExcelAux.iFormula = 0
            Case EXCEL_FORMULA_SUM_TEXTO
                objBrowseExcelAux.iFormula = EXCEL_FORMULA_SUM
            Case EXCEL_FORMULA_COUNT_TEXTO
                objBrowseExcelAux.iFormula = EXCEL_FORMULA_COUNT
            Case EXCEL_FORMULA_MIN_TEXTO
                objBrowseExcelAux.iFormula = EXCEL_FORMULA_MIN
            Case EXCEL_FORMULA_MAX_TEXTO
                objBrowseExcelAux.iFormula = EXCEL_FORMULA_MAX
            Case EXCEL_FORMULA_AVG_TEXTO
                objBrowseExcelAux.iFormula = EXCEL_FORMULA_AVG
        End Select
        Select Case GridTabCampos.TextMatrix(iIndice, COL_GRIDCAMPOS_POSICAO)
            Case ""
                objBrowseExcelAux.iPosicao = 0
            Case EXCEL_TABDIN_POS_LINHA_TEXTO
                objBrowseExcelAux.iPosicao = EXCEL_TABDIN_POS_LINHA
            Case EXCEL_TABDIN_POS_COLUNA_TEXTO
                objBrowseExcelAux.iPosicao = EXCEL_TABDIN_POS_COLUNA
            Case EXCEL_TABDIN_POS_FILTRO_TEXTO
                objBrowseExcelAux.iPosicao = EXCEL_TABDIN_POS_FILTRO
            Case EXCEL_TABDIN_POS_VALOR_TEXTO
                objBrowseExcelAux.iPosicao = EXCEL_TABDIN_POS_VALOR
        End Select
        If objBrowseExcelAux.iPosicao <> 0 Then
            objBrowseConfigura1.objBrowseExcel.colCampos.Add objBrowseExcelAux
        End If
    Next
    
    objBrowseConfigura1.objBrowseExcel.sArquivo = Arquivo.Text
    
    objBrowseConfigura1.iTelaOK = OK
    
    Unload BrowseConfigura
        
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr
    
        Case 9123
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BROWSE_SEM_COLUNAS1", gErr)

        Case 20657
        
        Case 202232
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCALIZACAO_NAO_PREENCHIDA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143885)
        
    End Select

    Exit Sub
        
End Sub

Private Sub BotaoParaBaixo_Click()

Dim sNome As String
Dim iIndice As Integer

    If CamposPosicionados.ListCount > 1 And CamposPosicionados.ListIndex < CamposPosicionados.ListCount - 1 Then
        iIndice = CamposPosicionados.ListIndex
        sNome = CamposPosicionados.List(CamposPosicionados.ListIndex + 1)
        CamposPosicionados.List(CamposPosicionados.ListIndex + 1) = CamposPosicionados.List(CamposPosicionados.ListIndex)
        CamposPosicionados.List(CamposPosicionados.ListIndex) = sNome
        CamposPosicionados.ListIndex = CamposPosicionados.ListIndex + 1
        sNome = CamposExibidos.List(iIndice + 1)
        CamposExibidos.List(iIndice + 1) = CamposExibidos.List(iIndice)
        CamposExibidos.List(iIndice) = sNome
        sNome = TitulosCampos.List(iIndice + 1)
        TitulosCampos.List(iIndice + 1) = TitulosCampos.List(iIndice)
        TitulosCampos.List(iIndice) = sNome
    End If

End Sub

Private Sub BotaoParaBaixoOrdenacao_Click()

Dim sNome As String
Dim iIndice As Integer

    If ListaOrdenacao.ListCount > 1 And ListaOrdenacao.ListIndex < ListaOrdenacao.ListCount - 1 Then
        iIndice = ListaOrdenacao.ListIndex
        sNome = ListaOrdenacao.List(ListaOrdenacao.ListIndex + 1)
        ListaOrdenacao.List(ListaOrdenacao.ListIndex + 1) = ListaOrdenacao.List(ListaOrdenacao.ListIndex)
        ListaOrdenacao.List(ListaOrdenacao.ListIndex) = sNome
        ListaOrdenacao.ListIndex = ListaOrdenacao.ListIndex + 1
    End If

End Sub

Private Sub BotaoParaCima_Click()

Dim sNome As String
Dim iIndice As Integer

    If CamposPosicionados.ListCount > 1 And CamposPosicionados.ListIndex > 0 Then
        iIndice = CamposPosicionados.ListIndex
        sNome = CamposPosicionados.List(CamposPosicionados.ListIndex - 1)
        CamposPosicionados.List(CamposPosicionados.ListIndex - 1) = CamposPosicionados.List(CamposPosicionados.ListIndex)
        CamposPosicionados.List(CamposPosicionados.ListIndex) = sNome
        CamposPosicionados.ListIndex = CamposPosicionados.ListIndex - 1
        sNome = CamposExibidos.List(iIndice - 1)
        CamposExibidos.List(iIndice - 1) = CamposExibidos.List(iIndice)
        CamposExibidos.List(iIndice) = sNome
        sNome = TitulosCampos.List(iIndice - 1)
        TitulosCampos.List(iIndice - 1) = TitulosCampos.List(iIndice)
        TitulosCampos.List(iIndice) = sNome
    End If
    
End Sub

Private Sub BotaoParaCimaOrdenacao_Click()

Dim sNome As String
Dim iIndice As Integer

    If ListaOrdenacao.ListCount > 1 And ListaOrdenacao.ListIndex > 0 Then
        iIndice = ListaOrdenacao.ListIndex
        sNome = ListaOrdenacao.List(ListaOrdenacao.ListIndex - 1)
        ListaOrdenacao.List(ListaOrdenacao.ListIndex - 1) = ListaOrdenacao.List(ListaOrdenacao.ListIndex)
        ListaOrdenacao.List(ListaOrdenacao.ListIndex) = sNome
        ListaOrdenacao.ListIndex = ListaOrdenacao.ListIndex - 1
    End If

End Sub

Private Sub BotaoRemoverCampoOrdenacao_Click()

    If ListaOrdenacao.ListIndex <> -1 Then
    
        ListaCamposDisponiveis.AddItem ListaOrdenacao.Text
        ListaOrdenacao.RemoveItem (ListaOrdenacao.ListIndex)
    
    End If

End Sub

Private Sub BotaoRemoverOrdenacao_Click()

Dim objBrowseIndice As AdmBrowseIndice
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoRemoverOrdenacao_Click

    If Ordenacao.ListIndex = -1 Then Error 60831
    
    For iIndice = 1 To objBrowseConfigura1.colBrowseIndiceUsuario.Count

        Set objBrowseIndice = objBrowseConfigura1.colBrowseIndiceUsuario.Item(iIndice)
        
        If Ordenacao.ItemData(Ordenacao.ListIndex) = objBrowseIndice.iIndice Then
            objBrowseConfigura1.colBrowseIndiceUsuario.Remove (iIndice)
            Ordenacao.RemoveItem (Ordenacao.ListIndex)
            Exit For
        End If

    Next
    
    Call BotaoLimparOrdenacao_Click
    
    objBrowseConfigura1.iAlteradoOrdenacao = 1
    
    Exit Sub
    
Erro_BotaoRemoverOrdenacao_Click:

    Select Case Err
    
        Case 60831
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDENACAO_NAO_SELECIONADA", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143886)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoSelecionar_Click()

    If CamposDisponiveis.ListIndex > -1 Then
        If CamposDisponiveis.Selected(CamposDisponiveis.ListIndex) = False Then
            CamposDisponiveis.Selected(CamposDisponiveis.ListIndex) = True
        End If
    End If
    
End Sub

Private Sub BotaoSelecionarTodos_Click()

Dim iIndice As Integer
Dim iIndice1 As Integer

    For iIndice = 0 To CamposDisponiveis.ListCount - 1
        If CamposDisponiveis.Selected(iIndice) = False Then
            CamposDisponiveis.Selected(iIndice) = True
        End If
    Next
    
    CamposDisponiveis.ListIndex = -1
    
End Sub

Private Sub CamposDisponiveis_ItemCheck(Item As Integer)

Dim iIndice As Integer
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo

    If CamposDisponiveis.Selected(Item) = False Then
        For iIndice = 0 To CamposPosicionados.ListCount - 1
            If CamposPosicionados.List(iIndice) = CamposDisponiveis.List(Item) Then
                CamposPosicionados.RemoveItem (iIndice)
                CamposExibidos.RemoveItem (iIndice)
                TitulosCampos.RemoveItem (iIndice)
                Exit For
            End If
        Next
        'Call Trata_GridExcel(True, CamposDisponiveis.List(Item))
    Else
        CamposPosicionados.AddItem CamposDisponiveis.List(Item)
        CamposExibidos.AddItem CamposDisponiveis.List(Item)
        For Each objBrowseUsuarioCampo In colCamposDisponiveis
            If objBrowseUsuarioCampo.sNome = CamposDisponiveis.List(Item) Then
                TitulosCampos.AddItem objBrowseUsuarioCampo.sTitulo
                Exit For
            End If
        Next
        'Call Trata_GridExcel(False, CamposDisponiveis.List(Item))
    End If
    
End Sub

Private Sub CamposExibidos_Click()
    Titulo.Text = TitulosCampos.List(CamposExibidos.ListIndex)
End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objGrid1 = New AdmGrid
    Set objGridForm = New AdmGrid
    Set objGridCampos = New AdmGrid
    
    'inicializacao do grid
    Call Inicializa_Grid_Selecao
    Call Inicializa_Grid_Formulas
    Call Inicializa_Grid_Campos

    iFrameAtual = 0
    iFrame2Atual = 0

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143887)
        
    End Select

    Exit Sub
    
End Sub

Function Trata_Parametros(objBrowseConfigura As AdmBrowseConfigura) As Long

Dim lErro As Long
Dim sNomeTela As String
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
Dim objBrowseIndice As AdmBrowseIndice

On Error GoTo Erro_Trata_Parametros

    sNomeTela = objBrowseConfigura.sNomeTela
    Set colBrowseUsuarioCampo = objBrowseConfigura.colBrowseUsuarioCampo
    Set objBrowseConfigura1 = objBrowseConfigura
    Set colCamposDisponiveis = New Collection
    
    lErro = Carga_Campos_Disponiveis(sNomeTela, colCamposDisponiveis)
    If lErro <> SUCESSO Then Error 9028
    
    lErro = Carga_Campos_Selecionados(colBrowseUsuarioCampo)
    If lErro <> SUCESSO Then Error 9029
    
    lErro = Trata_Dados_Export_XLS(objBrowseConfigura.objBrowseExcel)
    If lErro <> SUCESSO Then Error 9029
    
    'carrega a combo de campos do grid de selecao
    For Each objBrowseUsuarioCampo In colCamposDisponiveis
        ComboCampo.AddItem objBrowseUsuarioCampo.sNome
        ComboCampo.ItemData(ComboCampo.NewIndex) = objBrowseUsuarioCampo.iTipo
    Next
    
    'carrega a  lista de campos disponiveis para ordenacao
    For Each objBrowseUsuarioCampo In colCamposDisponiveis
        ListaCamposDisponiveis.AddItem objBrowseUsuarioCampo.sNome
    Next
    
    lErro = Grid_SelecaoSQL1_Preenche(objBrowseConfigura.sSelecaoSQL1Usuario)
    If lErro <> SUCESSO Then Error 20647
    
    BrowseConfigura.Caption = BrowseConfigura.Caption & sNomeTela
    
    If objBrowseConfigura.iPesquisa = ADM_CONFIGURA_PESQUISA Then
        Opcoes.Tabs.Item(TAB_INDEX_PESQUISA).Selected = True
    End If
    
    For Each objBrowseIndice In objBrowseConfigura1.colBrowseIndiceUsuario
    
        Ordenacao.AddItem objBrowseIndice.sNomeIndice
        Ordenacao.ItemData(Ordenacao.NewIndex) = objBrowseIndice.iIndice
    
    Next
    
    Set objBrowseConfigura1.colBrowseIndice = objBrowseConfigura.colBrowseIndice
    
    iAlteradoOrdenacao = objBrowseConfigura.iAlteradoOrdenacao
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case 9028, 9029, 20647
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143888)
    
    End Select
    
    Exit Function

End Function

Function Carga_Campos_Disponiveis(sNomeTela As String, colCamposDisponiveis As Collection) As String

Dim lErro As Long
Dim sCodGrupo As String
Dim objGrupoBrowseCampo As AdmGrupoBrowseCampo
Dim colGrupoBrowseCampo As New Collection
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
Dim objCampo As AdmCampos
Dim sCodUsuario As String

On Error GoTo Erro_Carga_Campos_Disponiveis
    
    sCodUsuario = gsUsuario

'    'obtem o codigo do usuario
'    lErro = Obter_Usuario(sCodUsuario)
'    If lErro <> SUCESSO Then Error 9118

    sCodGrupo = String(STRING_GRUPO, 0)

    If giLocalOperacao = LOCALOPERACAO_ECF Then
        sCodGrupo = "supervisor"
    Else
    'obtem o codigo do grupo
    lErro = Obter_Grupo(sCodGrupo)
    If lErro <> SUCESSO Then Error 9114
    End If
    
    'le os campos disponiveis para a tela x grupo em questão
    lErro = CF("GrupoBrowseCampo_Le", sCodGrupo, sNomeTela, colGrupoBrowseCampo)
    If lErro <> SUCESSO Then Error 9115
    
    For Each objGrupoBrowseCampo In colGrupoBrowseCampo
    
        Set objCampo = New AdmCampos
            
        objCampo.sNomeArq = objGrupoBrowseCampo.sNomeArq
        objCampo.sNome = objGrupoBrowseCampo.sNome
        
        lErro = CF("Campos_Le", objCampo)
        If lErro <> SUCESSO And lErro <> 9061 Then Error 9116
        
        'se o campo não estiver cadastrado
        If lErro = 9061 Then Error 9117
    
        Set objBrowseUsuarioCampo = New AdmBrowseUsuarioCampo
            
        objBrowseUsuarioCampo.sNomeTela = sNomeTela
        objBrowseUsuarioCampo.sCodUsuario = sCodUsuario
        objBrowseUsuarioCampo.sNomeArq = objCampo.sNomeArq
        objBrowseUsuarioCampo.sNome = objCampo.sNome
        objBrowseUsuarioCampo.sTitulo = objCampo.sTituloGrid
        objBrowseUsuarioCampo.lLargura = 1000
        objBrowseUsuarioCampo.iTipo = objCampo.iTipo
            
        colCamposDisponiveis.Add objBrowseUsuarioCampo
        
        CamposDisponiveis.AddItem objBrowseUsuarioCampo.sNome
        
    Next
    
    Carga_Campos_Disponiveis = SUCESSO

    Exit Function

Erro_Carga_Campos_Disponiveis:

    Carga_Campos_Disponiveis = Err

    Select Case Err

        Case 9114
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OBTENCAO_CODIGO_GRUPO", Err)
        
        Case 9115, 9116
        
        Case 9117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAMPO_NAO_CADASTRADO", Err, objCampo.sNome, objCampo.sNomeArq)
        
        Case 9118
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OBTENCAO_CODIGO_USUARIO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143889)

    End Select

    Exit Function
    
End Function

Function Carga_Campos_Selecionados(colBrowseUsuarioCampo As Collection) As String

Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
Dim iIndice As Integer
Dim iAchou As Integer

    For Each objBrowseUsuarioCampo In colBrowseUsuarioCampo
    
        For iIndice = 0 To CamposDisponiveis.ListCount - 1
        
            If CamposDisponiveis.List(iIndice) = objBrowseUsuarioCampo.sNome Then
                CamposDisponiveis.Selected(iIndice) = True
                Exit For
            End If
            
        Next
        
        CamposPosicionados.RemoveItem (CamposPosicionados.NewIndex)
        CamposExibidos.RemoveItem (CamposExibidos.NewIndex)
        TitulosCampos.RemoveItem (TitulosCampos.NewIndex)
        
        iAchou = 0
        
        For iIndice = 0 To CamposPosicionados.ListCount - 1
            If objBrowseUsuarioCampo.iPosicaoTela < CamposPosicionados.ItemData(iIndice) Then
                CamposPosicionados.AddItem objBrowseUsuarioCampo.sNome, iIndice
                CamposPosicionados.ItemData(CamposPosicionados.NewIndex) = objBrowseUsuarioCampo.iPosicaoTela
                CamposExibidos.AddItem objBrowseUsuarioCampo.sNome, iIndice
                TitulosCampos.AddItem objBrowseUsuarioCampo.sTitulo, iIndice
                iAchou = 1
                Exit For
            End If
        Next
        
        If iAchou = 0 Then
            CamposPosicionados.AddItem objBrowseUsuarioCampo.sNome
            CamposPosicionados.ItemData(CamposPosicionados.NewIndex) = objBrowseUsuarioCampo.iPosicaoTela
            CamposExibidos.AddItem objBrowseUsuarioCampo.sNome
            TitulosCampos.AddItem objBrowseUsuarioCampo.sTitulo
        End If
        
    Next
    
    Carga_Campos_Selecionados = SUCESSO
    
End Function

Private Sub FormatoCSV_Click()
    Call Trata_Formato
End Sub

Private Sub FormatoXls_Click()
    Call Trata_Formato
End Sub

Private Sub IncluirGrafico_Click()
    If IncluirGrafico.Value = vbChecked Then
        TipoGrafico.Enabled = True
        TipoGrafico.ListIndex = 0
    Else
        TipoGrafico.Enabled = False
        TipoGrafico.ListIndex = -1
    End If
End Sub

Private Sub IncluiTabela_Click()
Dim iLinha As Integer
    If IncluiTabela.Value = vbChecked Then
        FrameGraf.Enabled = True
        FrameCampos.Enabled = True
    Else
        FrameGraf.Enabled = False
        IncluirGrafico.Value = vbUnchecked
        TipoGrafico.ListIndex = -1
        FrameCampos.Enabled = False
        Call Grid_Limpa(objGridCampos)
'        For iLinha = 1 To objGridCampos.iLinhasExistentes
'            GridTabCampos.TextMatrix(iLinha, COL_GRIDCAMPOS_POSICAO) = ""
'            GridTabCampos.TextMatrix(iLinha, COL_GRIDCAMPOS_FORMULA) = ""
'        Next
    End If
End Sub

Private Sub ListaCamposDisponiveis_DblClick()
    Call BotaoInserirCampoOrdenacao_Click
End Sub

Private Sub ListaOrdenacao_DblClick()
    Call BotaoRemoverCampoOrdenacao_Click
End Sub

Private Sub Opcoes_Click()

    If Opcoes.SelectedItem.Index - 1 <> iFrameAtual Then
        Frame1(Opcoes.SelectedItem.Index - 1).Visible = True
        Frame1(iFrameAtual).Visible = False
        iFrameAtual = Opcoes.SelectedItem.Index - 1
    End If
    
    If Opcoes.Tabs.Item(iFrameAtual + 1).Caption = "Títulos" Then
        If CamposExibidos.ListIndex >= 0 Then
            Titulo.Text = TitulosCampos.List(CamposExibidos.ListIndex)
        Else
            Titulo.Text = ""
        End If
    End If
    
    
    
End Sub

Private Sub Ordenacao_Click()
    
Dim objBrowseIndice As AdmBrowseIndice
Dim colBrowseIndiceSegmentos As New Collection
Dim objBrowseIndiceSegmentos As AdmBrowseIndiceSegmentos
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
Dim iAchou As Integer
Dim iIndice As Integer

    If Ordenacao.ListIndex <> -1 Then

        Call CF("BrowseIndiceSegmentos_Le", objBrowseConfigura1.sNomeTela, Ordenacao.ItemData(Ordenacao.ListIndex), colBrowseIndiceSegmentos, objBrowseConfigura1.colBrowseIndiceUsuario)
        
        ListaOrdenacao.Clear
        
        For iIndice = 1 To colBrowseIndiceSegmentos.Count
            For Each objBrowseIndiceSegmentos In colBrowseIndiceSegmentos
                If objBrowseIndiceSegmentos.iPosicaoCampo = iIndice Then Exit For
            Next
                
            ListaOrdenacao.AddItem objBrowseIndiceSegmentos.sNomeCampo
        Next
        
        ListaCamposDisponiveis.Clear
        
        'carrega a combo de campos da lista de campos disponiveis para ordenacao
        For Each objBrowseUsuarioCampo In colCamposDisponiveis
            
            iAchou = 0
            
            For Each objBrowseIndiceSegmentos In colBrowseIndiceSegmentos
                If objBrowseIndiceSegmentos.sNomeCampo = objBrowseUsuarioCampo.sNome Then
                    iAchou = 1
                    Exit For
                End If
            Next
            
            If iAchou = 0 Then ListaCamposDisponiveis.AddItem objBrowseUsuarioCampo.sNome
                            
        Next
    
    End If
    
End Sub

Private Sub Titulo_Change()
    If CamposExibidos.ListIndex >= 0 Then
        TitulosCampos.List(CamposExibidos.ListIndex) = Titulo.Text
    End If
End Sub

Private Function Grid_SelecaoSQL1_Preenche(sSelecaoSQL1Usuario As String) As Long

Dim sSQL As String
Dim sToken As String
Dim lPosicao As Long
Dim iPosToken As Integer
Dim iLinha As Long
Dim lErro As Long
Dim iVersao As Integer

On Error GoTo Erro_Grid_SelecaoSQL1_Preenche

    iLinha = 1
    
    If left(sSelecaoSQL1Usuario, 1) = "*" Then
        iVersao = 1
        sSQL = Mid(sSelecaoSQL1Usuario, 3)
    Else
        sSQL = Trim(sSelecaoSQL1Usuario)
    End If
    
    
    Do While Len(sSQL) > 0
    
        If iPosToken = objGrid1.colCampo.Count Then
            iPosToken = 1
            iLinha = iLinha + 1
        Else
            iPosToken = iPosToken + 1
        End If
    
        If iPosToken = COL_GRIDSELECAO_PAR_ABRIR Then
            objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
        End If
    
        If iVersao = 0 And (iPosToken = COL_GRIDSELECAO_PAR_FECHAR Or iPosToken = COL_GRIDSELECAO_PAR_ABRIR) Then iPosToken = iPosToken + 1
    
        'no caso dos valores eles são delimitados por aspas
        If iPosToken = COL_GRIDSELECAO_VALOR Then
            sSQL = Mid(sSQL, 2)
            lPosicao = InStr(sSQL, Chr(34))
        Else
            lPosicao = InStr(sSQL, " ")
        End If
        
        If lPosicao = 0 Then lPosicao = Len(sSQL) + 1
    
        sToken = left(sSQL, lPosicao - 1)
        
'        If iVersao = 1 And sToken = "" And (iPosToken = COL_GRIDSELECAO_PAR_ABRIR Or iPosToken = COL_GRIDSELECAO_PAR_FECHAR) Then
'            lPosicao = lPosicao + 1
'        End If
        
        'no caso dos valores eles são delimitados por aspas
        If iPosToken = COL_GRIDSELECAO_VALOR Then
            sSQL = Mid(sSQL, lPosicao + 2)
        Else
            sSQL = Mid(sSQL, lPosicao + 1)
        End If
        
        GridSelecao.TextMatrix(iLinha, iPosToken) = sToken
            
        If iPosToken = COL_GRIDSELECAO_PAR_ABRIR Then sSQL = Trim(sSQL)
        
    Loop

    Grid_SelecaoSQL1_Preenche = SUCESSO
    
    Exit Function
    
Erro_Grid_SelecaoSQL1_Preenche:

    Grid_SelecaoSQL1_Preenche = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143890)

    End Select

    Exit Function

End Function

Function Inicializa_Grid_Selecao() As Long
   
Dim lErro As Long
   
On Error GoTo Erro_Inicializa_Grid_Selecao
   
    'tela em questão
    Set objGrid1.objForm = Me
    
    'titulos do grid
    objGrid1.colColuna.Add ("")
    objGrid1.colColuna.Add ("(")
    objGrid1.colColuna.Add ("Campo")
    objGrid1.colColuna.Add ("Operador")
    objGrid1.colColuna.Add ("Valor")
    objGrid1.colColuna.Add (")")
    objGrid1.colColuna.Add ("E/OU")
    
   'campos de edição do grid
    objGrid1.colCampo.Add (ComboParAbrir.Name)
    objGrid1.colCampo.Add (ComboCampo.Name)
    objGrid1.colCampo.Add (ComboOperacao.Name)
    objGrid1.colCampo.Add (Valor.Name)
    objGrid1.colCampo.Add (ComboParFechar.Name)
    objGrid1.colCampo.Add (ComboConjuncao.Name)
    
    objGrid1.objGrid = GridSelecao
   
    'todas as linhas do grid
    objGrid1.objGrid.Rows = 11
    
    'linhas visiveis do grid sem contar com as linhas fixas
    objGrid1.iLinhasVisiveis = 5
    
    objGrid1.objGrid.ColWidth(0) = 300
    
    objGrid1.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGrid1)
    
    Inicializa_Grid_Selecao = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_Selecao:

    Inicializa_Grid_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143891)

    End Select

    Exit Function
    
End Function

Private Sub ComboCampo_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub ComboCampo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub ComboCampo_LostFocus()

    Set objGrid1.objControle = ComboCampo
    Call Grid_Campo_Libera_Foco_Modal(objGrid1)

End Sub

Private Sub ComboOperacao_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub ComboOperacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub ComboOperacao_LostFocus()

    Set objGrid1.objControle = ComboOperacao
    Call Grid_Campo_Libera_Foco_Modal(objGrid1)

End Sub

Private Sub ComboConjuncao_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub ComboConjuncao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub ComboConjuncao_LostFocus()

    Set objGrid1.objControle = ComboConjuncao
    Call Grid_Campo_Libera_Foco_Modal(objGrid1)

End Sub

Private Sub ComboParAbrir_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub ComboParAbrir_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub ComboParAbrir_LostFocus()

    Set objGrid1.objControle = ComboParAbrir
    Call Grid_Campo_Libera_Foco_Modal(objGrid1)

End Sub

Private Sub ComboParFechar_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub ComboParFechar_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub ComboParFechar_LostFocus()

    Set objGrid1.objControle = ComboParFechar
    Call Grid_Campo_Libera_Foco_Modal(objGrid1)

End Sub

Private Sub Valor_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Valor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub Valor_LostFocus()

    Set objGrid1.objControle = Valor
    Call Grid_Campo_Libera_Foco_Modal(objGrid1)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then
    
        'Verifica qual é o grid
        If objGridInt.objGrid.Name = GridSelecao.Name Then
    
            Select Case GridSelecao.Col
        
                Case COL_GRIDSELECAO_PAR_ABRIR
                
                    lErro = Saida_Celula_Par_Abrir(objGridInt)
                    If lErro <> SUCESSO Then gError 178397
        
                Case COL_GRIDSELECAO_CAMPO
                
                    lErro = Saida_Celula_Campo(objGridInt)
                    If lErro <> SUCESSO Then gError 20648
                    
                Case COL_GRIDSELECAO_OP
                
                    lErro = Saida_Celula_Op(objGridInt)
                    If lErro <> SUCESSO Then gError 20649
                    
                Case COL_GRIDSELECAO_VALOR
                
                    lErro = Saida_Celula_Valor(objGridInt)
                    If lErro <> SUCESSO Then gError 20650
                    
                Case COL_GRIDSELECAO_PAR_FECHAR
                
                    lErro = Saida_Celula_Par_Fechar(objGridInt)
                    If lErro <> SUCESSO Then gError 178398
                    
                Case COL_GRIDSELECAO_EOU
                
                    lErro = Saida_Celula_EOU(objGridInt)
                    If lErro <> SUCESSO Then gError 20651
    
            End Select
        ElseIf objGridInt.objGrid.Name = GridTabCampos.Name Then

            Select Case GridTabCampos.Col
                            
                Case COL_GRIDCAMPOS_CAMPO
                
                    lErro = Saida_Celula_Padrao(objGridInt, TabCamposCampos, True, True)
                    If lErro <> SUCESSO Then gError 178397
                            
                Case COL_GRIDCAMPOS_FORMULA
                
                    lErro = Saida_Celula_Padrao(objGridInt, TabCamposForm)
                    If lErro <> SUCESSO Then gError 178397
                    
                Case COL_GRIDCAMPOS_POSICAO
                
                    lErro = Saida_Celula_Padrao(objGridInt, TabCamposPosicao)
                    If lErro <> SUCESSO Then gError 178397
                    
            End Select

        ElseIf objGridInt.objGrid.Name = GridFormulas.Name Then
        
            Select Case GridFormulas.Col
                            
                Case COL_GRIDFORMULAS_CAMPO
                
                    lErro = Saida_Celula_Padrao(objGridInt, FormCampos, True, True)
                    If lErro <> SUCESSO Then gError 178397
                    
                Case COL_GRIDFORMULAS_FORMULA
                
                    lErro = Saida_Celula_Padrao(objGridInt, FormFormulas)
                    If lErro <> SUCESSO Then gError 178397
                    
            End Select
            
        End If
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 20652
        
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula:

    Saida_Celula = gErr
    
    Select Case gErr
    
        Case 20648, 20649, 20650, 20651, 178397, 178398
        
        Case 20652
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143892)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Campo(objGridInt As AdmGrid) As Long
'faz a critica da celula preenchimento do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Campo

    Set objGridInt.objControle = ComboCampo
                
    If Len(Trim(ComboCampo.Text)) > 0 And GridSelecao.Row - GridSelecao.FixedRows = objGridInt.iLinhasExistentes Then
       objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 20653

    Saida_Celula_Campo = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Campo:

    Saida_Celula_Campo = Err
    
    Select Case Err
    
        Case 20653
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143893)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Op(objGridInt As AdmGrid) As Long
'faz a critica da celula preenchimento do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Op

    Set objGridInt.objControle = ComboOperacao
                
    If Len(Trim(ComboOperacao.Text)) > 0 And GridSelecao.Row - GridSelecao.FixedRows = objGridInt.iLinhasExistentes Then
       objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 20654

    Saida_Celula_Op = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Op:

    Saida_Celula_Op = Err
    
    Select Case Err
    
        Case 20654
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143894)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_EOU(objGridInt As AdmGrid) As Long
'faz a critica da celula preenchimento do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_EOU

    Set objGridInt.objControle = ComboConjuncao
                
    If Len(Trim(ComboConjuncao.Text)) > 0 And GridSelecao.Row - GridSelecao.FixedRows = objGridInt.iLinhasExistentes Then
       objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 20655

    Saida_Celula_EOU = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_EOU:

    Saida_Celula_EOU = Err
    
    Select Case Err
    
        Case 20655
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143895)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'faz a critica da celula preenchimento do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = Valor
                
    If Len(Trim(Valor.Text)) > 0 Then
    
        'o valor não pode conter aspas no seu interior
        If InStr(Valor.Text, Chr(34)) <> 0 Then Error 20672
    
        If GridSelecao.Row - GridSelecao.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 20656

    Saida_Celula_Valor = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = Err
    
    Select Case Err
    
        Case 20672
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAMPO_NAO_PODE_CONTER_ASPAS", Err, Error$)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 20656
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143896)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Par_Abrir(objGridInt As AdmGrid) As Long
'faz a critica da celula preenchimento do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Par_Abrir

    Set objGridInt.objControle = ComboParAbrir
                
    If Len(Trim(ComboParAbrir.Text)) > 0 And GridSelecao.Row - GridSelecao.FixedRows = objGridInt.iLinhasExistentes Then
       objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 178399

    Saida_Celula_Par_Abrir = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Par_Abrir:

    Saida_Celula_Par_Abrir = gErr
    
    Select Case gErr
    
        Case 178399
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178400)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Par_Fechar(objGridInt As AdmGrid) As Long
'faz a critica da celula preenchimento do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Par_Fechar

    Set objGridInt.objControle = ComboParFechar
                
    If Len(Trim(ComboParFechar.Text)) > 0 And GridSelecao.Row - GridSelecao.FixedRows = objGridInt.iLinhasExistentes Then
       objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 178401

    Saida_Celula_Par_Fechar = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Par_Fechar:

    Saida_Celula_Par_Fechar = gErr
    
    Select Case gErr
    
        Case 178401
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178402)
        
    End Select

    Exit Function

End Function

Private Function Move_GridSelecao_Memoria() As Long

Dim sSQLUsuario As String
Dim sSQL As String
Dim iLinha As Integer
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
Dim iPosToken As Integer
Dim lErro As Long
Dim sAux As String
Dim iPos As Integer
Dim iPos1 As Integer
Dim sAux1 As String
Dim sAux2 As String
Dim colMnemonico As New Collection
Dim iInicio As Integer
Dim iTamanho As Integer
Dim iPosTransferido As Integer
Dim iParAbertos As Integer

On Error GoTo Erro_Move_GridSelecao_Memoria

    sSQLUsuario = "*"
    sSQL = ""

    For iLinha = 1 To objGrid1.iLinhasExistentes
    
        For iPosToken = 1 To objGrid1.colCampo.Count
        
            If iPosToken = COL_GRIDSELECAO_CAMPO Then
        
                If Len(GridSelecao.TextMatrix(iLinha, iPosToken)) = 0 Then gError 20663
        
                For Each objBrowseUsuarioCampo In colCamposDisponiveis
            
                    If objBrowseUsuarioCampo.sNome = GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_CAMPO) Then Exit For
                    
                Next
                
                sSQLUsuario = sSQLUsuario & " " & GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_CAMPO)
                sSQL = sSQL & " " & GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_CAMPO)
                
            ElseIf iPosToken = COL_GRIDSELECAO_VALOR Then
            
                If Len(GridSelecao.TextMatrix(iLinha, iPosToken)) = 0 Then gError 20664
            
                Select Case objBrowseUsuarioCampo.iTipo
                
                    Case ADM_TIPO_SMALLINT
                        lErro = Valor_Inteiro_Critica(GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_VALOR))
                        If lErro <> SUCESSO Then gError 20658
                        
                        sSQL = sSQL & " " & CStr(CInt(GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_VALOR)))
                        
                    Case ADM_TIPO_INTEGER
                        lErro = Valor_Long_Critica(GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_VALOR))
                        If lErro <> SUCESSO Then gError 20659
                        
                        sSQL = sSQL & " " & CStr(CLng(GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_VALOR)))
                        
                    Case ADM_TIPO_DOUBLE
                        lErro = Valor_Double_Critica(GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_VALOR))
                        If lErro <> SUCESSO Then gError 20660
                        
                        sAux = Format(GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_VALOR), "General Number")
                        iPos = InStr(sAux, ",")
                        If iPos <> 0 Then sAux = Mid(sAux, 1, iPos - 1) + "." + Mid(sAux, iPos + 1)
                        sSQL = sSQL & " " & sAux
                        
                    Case ADM_TIPO_DATE
'                        lErro = Valor_Date_Critica(GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_VALOR))
'                        If lErro <> SUCESSO Then gError 20661
'
'                        sSQL = sSQL & " {d" & Format(CDate(GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_VALOR)), "'yyyy-mm-dd'") & " }"
                    
                        sAux = GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_VALOR)
                        
                        iPos = 1
                        iPosTransferido = 0
                        sAux2 = ""
                        sAux1 = ""
                        
                        Do While iPos > 0
                        
                            iPos = InStr(iPos, sAux, "/")
                        
                            If iPos > 0 Then
                            
                                iPos1 = InStr(iPos + 1, sAux, "/")
                                
                                If iPos1 > 0 Then
                                                        
                                    If iPos1 - iPos = 2 Or iPos1 - iPos = 3 Then
                                    
                                        If IsNumeric(Mid(sAux, iPos + 1, iPos1 - (iPos + 1))) Then
                                                        
                                            If iPos1 + 1 <= Len(sAux) Then
                                        
                                                Do While IsNumeric(Mid(sAux, iPos1 + 1, 1))
                                                    iPos1 = iPos1 + 1
                                                    If iPos1 + 1 > Len(sAux) Then Exit Do
                                                Loop
                                            
                                            End If
                                            
                                            If iPos - 1 > 0 Then
                                                Do While IsNumeric(Mid(sAux, iPos - 1, 1))
                                                    iPos = iPos - 1
                                                    If iPos - 1 < 1 Then Exit Do
                                                Loop
                                    
                                            End If
                                    
                                            lErro = Valor_Date_Critica(Mid(sAux, iPos, (iPos1 - iPos) + 1))
                                            If lErro <> SUCESSO Then gError 178266
                                    
                                            'transfere a parte anterior a data
                                            If iPos - (iPosTransferido + 1) > 0 Then
                                                sAux1 = sAux1 & Mid(sAux, iPosTransferido + 1, iPos - (iPosTransferido + 1))
                                                sAux2 = sAux2 & Mid(sAux, iPosTransferido + 1, iPos - (iPosTransferido + 1))
                                            End If
                                            
                                            sAux1 = sAux1 & " {d" & Format(CDate(Mid(sAux, iPos, (iPos1 - iPos) + 1)), "'yyyy-mm-dd'") & " } "
                                            sAux2 = sAux2 & "'" & Format(Mid(sAux, iPos, (iPos1 - iPos) + 1), "dd/mm/yyyy") & "'"
                                            
                                            iPos = iPos1 + 1
                                            iPosTransferido = iPos1
                            
                                        End If
                            
                                    End If
                            
                                Else
                                
                                    gError 178451
                                
                                End If
                            
                            End If
                            
                        Loop
                    
                        If Len(sAux) - iPosTransferido > 0 Then
                            sAux1 = sAux1 & Mid(sAux, iPosTransferido + 1, Len(sAux) - iPosTransferido)
                            sAux2 = sAux2 & Mid(sAux, iPosTransferido + 1, Len(sAux) - iPosTransferido)
                        End If
                        
                        lErro = CF("Valida_Formula_Browser", sAux2, TIPO_DATA, iInicio, iTamanho, colMnemonico)
                        If lErro <> SUCESSO Then gError 178265
                    
                        iPos = 1
                        iPosTransferido = 0
                        
                        sAux = sAux1
                        sAux1 = ""
                        
                        Do While iPos > 0
                        
                            iPos = InStr(iPos, sAux, "DATA_HOJE()")
                        
                            If iPos > 0 Then
                        
                                                
                                If iPos - (iPosTransferido + 1) Then
                                    sAux1 = sAux1 & Mid(sAux, iPosTransferido + 1, iPos - (iPosTransferido + 1))
                                End If
                                
                                sAux1 = sAux1 & "{fn CURDATE()}"
                                
                                iPos = iPos + Len("DATA_HOJE()")
                                iPosTransferido = iPos - 1
                                
                            End If
                            
                        Loop
                    
                        
                        If Len(sAux) - iPosTransferido > 0 Then
                            sAux = sAux1 & Mid(sAux, iPosTransferido + 1, Len(sAux) - iPosTransferido)
                        Else
                            sAux = sAux1
                        End If
                        
                        If IsNumeric(Mid(sAux, 5, 4)) Then
                            If StrParaLong(Mid(sAux, 5, 4)) < 1900 Then Error 20661
                        End If
                            
                        sSQL = sSQL & sAux
                    
                    
                    Case ADM_TIPO_VARCHAR
                    
                        sAux = Replace(GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_VALOR), "''", vbNewLine)
                        sAux = Replace(sAux, "'", "''")
                        sAux = Replace(sAux, vbNewLine, "''")
                    
                        If GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_OP) = OP_LIKE Then
                            sSQL = sSQL & " '" & sAux & "%'"
                        Else
                            sSQL = sSQL & " '" & sAux & "'"
                        End If
                    
                    Case Else
                        Error 20662
                        
                End Select
            
                'coloca aspas entre qualquer valor para servir de delimitador em vez dos espaços. Isto permitirá inserir espaços como valor
                sSQLUsuario = sSQLUsuario & " " & Chr(34) & GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_VALOR) & Chr(34)
            
            ElseIf iPosToken = COL_GRIDSELECAO_EOU Then
            
                If iLinha < objGrid1.iLinhasExistentes Then
            
                    If Len(GridSelecao.TextMatrix(iLinha, iPosToken)) = 0 And objGrid1.iLinhasExistentes > iLinha Then Error 20666
            
                    If GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_EOU) = "E" Then
                        sSQLUsuario = sSQLUsuario & " E"
                        sSQL = sSQL & " AND"
                    End If
        
                    If GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_EOU) = "OU" Then
                        sSQLUsuario = sSQLUsuario & " OU"
                        sSQL = sSQL & " OR"
                    End If
                    
                End If
        
            ElseIf iPosToken = COL_GRIDSELECAO_OP Then
            
                If Len(GridSelecao.TextMatrix(iLinha, iPosToken)) = 0 Then Error 20671
            
                If GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_OP) = OP_LIKE And objBrowseUsuarioCampo.iTipo <> ADM_TIPO_VARCHAR Then Error 20670
        
                sSQLUsuario = sSQLUsuario & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
                sSQL = sSQL & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
        
            ElseIf iPosToken = COL_GRIDSELECAO_PAR_ABRIR Then
            
                Select Case GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_PAR_ABRIR)
                
                    Case "("
                        iParAbertos = iParAbertos + 1
                        sSQLUsuario = sSQLUsuario & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
                        sSQL = sSQL & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
        
                
                    Case "(("
                        iParAbertos = iParAbertos + 2
                        sSQLUsuario = sSQLUsuario & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
                        sSQL = sSQL & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
        
                
                    Case "((("
                        iParAbertos = iParAbertos + 3
                        sSQLUsuario = sSQLUsuario & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
                        sSQL = sSQL & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
        
                
                    Case "(((("
                        iParAbertos = iParAbertos + 4
                        sSQLUsuario = sSQLUsuario & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
                        sSQL = sSQL & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
        
                        
                    Case "((((("
                        iParAbertos = iParAbertos + 5
                        sSQLUsuario = sSQLUsuario & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
                        sSQL = sSQL & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
        
                    Case Else
                        sSQLUsuario = sSQLUsuario & " " & Trim(GridSelecao.TextMatrix(iLinha, iPosToken))
                        sSQL = sSQL & " " & Trim(GridSelecao.TextMatrix(iLinha, iPosToken))
                        
                End Select
        
            ElseIf iPosToken = COL_GRIDSELECAO_PAR_FECHAR Then
            
                Select Case GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_PAR_FECHAR)
                
                    Case ")"
                        iParAbertos = iParAbertos - 1
                        sSQLUsuario = sSQLUsuario & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
                        sSQL = sSQL & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
                
                    Case "))"
                        iParAbertos = iParAbertos - 2
                        sSQLUsuario = sSQLUsuario & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
                        sSQL = sSQL & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
                
                    Case ")))"
                        iParAbertos = iParAbertos - 3
                        sSQLUsuario = sSQLUsuario & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
                        sSQL = sSQL & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
                
                    Case "))))"
                        iParAbertos = iParAbertos - 4
                        sSQLUsuario = sSQLUsuario & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
                        sSQL = sSQL & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
                        
                    Case ")))))"
                        iParAbertos = iParAbertos - 5
                        sSQLUsuario = sSQLUsuario & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
                        sSQL = sSQL & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
                    
                    Case Else
                        sSQLUsuario = sSQLUsuario & " " & Trim(GridSelecao.TextMatrix(iLinha, iPosToken))
                        sSQL = sSQL & " " & Trim(GridSelecao.TextMatrix(iLinha, iPosToken))
                        
                        
                End Select
        
            Else
            
                If Len(GridSelecao.TextMatrix(iLinha, iPosToken)) = 0 Then Error 20665
            
                sSQLUsuario = sSQLUsuario & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
                sSQL = sSQL & " " & GridSelecao.TextMatrix(iLinha, iPosToken)
                
            End If
            
        Next
    
    Next

    If iParAbertos > 0 Then gError 178403

    If iParAbertos < 0 Then gError 178404

    objBrowseConfigura1.sSelecaoSQL1Usuario = sSQLUsuario
    objBrowseConfigura1.sSelecaoSQL1 = sSQL

    Move_GridSelecao_Memoria = SUCESSO
    
    Exit Function

Erro_Move_GridSelecao_Memoria:

    Move_GridSelecao_Memoria = gErr

    Select Case gErr
    
        Case 20658
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRIDSELECAO_INTEIRO_INVALIDO", gErr, iLinha, iPosToken, GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_VALOR))
            Opcoes.Tabs.Item(TAB_INDEX_PESQUISA).Selected = True
    
        Case 20659
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRIDSELECAO_LONG_INVALIDO", gErr, iLinha, iPosToken, GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_VALOR))
            Opcoes.Tabs.Item(TAB_INDEX_PESQUISA).Selected = True
    
        Case 20660
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRIDSELECAO_DOUBLE_INVALIDO", gErr, iLinha, iPosToken, GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_VALOR))
            Opcoes.Tabs.Item(TAB_INDEX_PESQUISA).Selected = True
    
        Case 20661
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRIDSELECAO_DATA_INVALIDA", gErr, iLinha, iPosToken, GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_VALOR))
            Opcoes.Tabs.Item(TAB_INDEX_PESQUISA).Selected = True
    
        Case 20662
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_CAMPO_INVALIDO1", gErr, iLinha, objBrowseUsuarioCampo.iTipo)
            Opcoes.Tabs.Item(TAB_INDEX_PESQUISA).Selected = True
    
        Case 20663, 20664, 20665, 20666, 20671
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRIDSELECAO_SEM_PREENCHIMENTO", gErr, iLinha, iPosToken)
            Opcoes.Tabs.Item(TAB_INDEX_PESQUISA).Selected = True
    
        Case 20670
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OPERADOR_LIKE", gErr, iLinha)
            Opcoes.Tabs.Item(TAB_INDEX_PESQUISA).Selected = True
    
        Case 178265
            Opcoes.Tabs.Item(TAB_INDEX_PESQUISA).Selected = True
    
        Case 178266, 178451
            Call Rotina_Erro(vbOKOnly, "ERRO_GRIDSELECAO_DATA_INVALIDA", gErr, iLinha, iPosToken, GridSelecao.TextMatrix(iLinha, COL_GRIDSELECAO_VALOR))
            Opcoes.Tabs.Item(TAB_INDEX_PESQUISA).Selected = True
    
        Case 178403
            Call Rotina_Erro(vbOKOnly, "ERRO_PAR_ABERTOS_SUPERA_FECHADOS", gErr, iParAbertos)
    
        Case 178404
            Call Rotina_Erro(vbOKOnly, "ERRO_PAR_FECHADOS_SUPERA_ABERTOS", gErr, Abs(iParAbertos))
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143897)
    
    End Select
    
    Exit Function
    
End Function

Private Sub GridSelecao_Click()

Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If
    
End Sub

Private Sub GridSelecao_GotFocus()
    
    Call Grid_Recebe_Foco(objGrid1)

End Sub

Private Sub GridSelecao_EnterCell()
    
    Call Grid_Entrada_Celula(objGrid1, iAlterado)
    
End Sub

Private Sub GridSelecao_LeaveCell()
    
    Call Saida_Celula(objGrid1)
    
End Sub

Private Sub GridSelecao_KeyDown(KeyCode As Integer, Shift As Integer)

Dim dColunaSoma As Double
Dim lErro As Long

On Error GoTo Erro_GridLancamentos_KeyDown

    lErro = Grid_Trata_Tecla1(KeyCode, objGrid1)
    If lErro <> SUCESSO Then Error 44124
    
    Exit Sub
    
Erro_GridLancamentos_KeyDown:

    Select Case Err
    
        Case 44124
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143898)
    
    End Select

    Exit Sub

End Sub

Private Sub GridSelecao_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridSelecao_LostFocus()
    
    Call Grid_Libera_Foco(objGrid1)

End Sub

Private Sub GridSelecao_RowColChange()

    Call Grid_RowColChange(objGrid1)
       
End Sub

Private Sub GridSelecao_Scroll()

    Call Grid_Scroll(objGrid1)
    
End Sub


Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
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

Private Sub Opcoes_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcoes)
End Sub

Function Inicializa_Grid_Formulas() As Long
   
Dim lErro As Long
   
On Error GoTo Erro_Inicializa_Grid_Formulas
   
    'tela em questão
    Set objGridForm.objForm = Me
    
    'titulos do grid
    objGridForm.colColuna.Add ("")
    objGridForm.colColuna.Add ("Campo")
    objGridForm.colColuna.Add ("Fórmula")
    
   'campos de edição do grid
    objGridForm.colCampo.Add (FormCampos.Name)
    objGridForm.colCampo.Add (FormFormulas.Name)
    
    objGridForm.objGrid = GridFormulas
   
    'todas as linhas do grid
    objGridForm.objGrid.Rows = 200 + 1
    
    'linhas visiveis do grid sem contar com as linhas fixas
    objGridForm.iLinhasVisiveis = 3
    
    objGridForm.objGrid.ColWidth(0) = 300
    
    objGridCampos.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    
    objGridForm.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGridForm)
    
    Inicializa_Grid_Formulas = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_Formulas:

    Inicializa_Grid_Formulas = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143891)

    End Select

    Exit Function
    
End Function

Function Inicializa_Grid_Campos() As Long
   
Dim lErro As Long
   
On Error GoTo Erro_Inicializa_Grid_Campos
   
    'tela em questão
    Set objGridCampos.objForm = Me
    
    'titulos do grid
    objGridCampos.colColuna.Add ("")
    objGridCampos.colColuna.Add ("Campo")
    objGridCampos.colColuna.Add ("Posição")
    objGridCampos.colColuna.Add ("Fórmula")
    
   'campos de edição do grid
    objGridCampos.colCampo.Add (TabCamposCampos.Name)
    objGridCampos.colCampo.Add (TabCamposPosicao.Name)
    objGridCampos.colCampo.Add (TabCamposForm.Name)
    
    objGridCampos.objGrid = GridTabCampos
   
    'todas as linhas do grid
    objGridCampos.objGrid.Rows = 200 + 1
    
    'linhas visiveis do grid sem contar com as linhas fixas
    objGridCampos.iLinhasVisiveis = 4
       
    objGridCampos.objGrid.ColWidth(0) = 300
    
    objGridCampos.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    
    objGridCampos.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGridCampos)
    
    Inicializa_Grid_Campos = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_Campos:

    Inicializa_Grid_Campos = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143891)

    End Select

    Exit Function
    
End Function

Private Sub FormCampos_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridForm)
End Sub

Private Sub FormCampos_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridForm)
End Sub

Private Sub FormCampos_LostFocus()
    Set objGrid1.objControle = FormCampos
    Call Grid_Campo_Libera_Foco_Modal(objGridForm)
End Sub

Private Sub FormFormulas_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridForm)
End Sub

Private Sub FormFormulas_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridForm)
End Sub

Private Sub FormFormulas_LostFocus()
    Set objGrid1.objControle = FormFormulas
    Call Grid_Campo_Libera_Foco_Modal(objGridForm)
End Sub

Private Sub TabCamposCampos_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCampos)
End Sub

Private Sub TabCamposCampos_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCampos)
End Sub

Private Sub TabCamposCampos_LostFocus()
    Set objGrid1.objControle = TabCamposCampos
    Call Grid_Campo_Libera_Foco_Modal(objGridCampos)
End Sub

Private Sub TabCamposForm_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCampos)
End Sub

Private Sub TabCamposForm_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCampos)
End Sub

Private Sub TabCamposForm_LostFocus()
    Set objGrid1.objControle = TabCamposForm
    Call Grid_Campo_Libera_Foco_Modal(objGridCampos)
End Sub

Private Sub TabCamposPosicao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCampos)
End Sub

Private Sub TabCamposPosicao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCampos)
End Sub

Private Sub TabCamposPosicao_LostFocus()
    Set objGrid1.objControle = TabCamposPosicao
    Call Grid_Campo_Libera_Foco_Modal(objGridCampos)
End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Private Sub TabStrip1_Click()

    If TabStrip1.SelectedItem.Index - 1 <> iFrame2Atual Then
        Frame2(TabStrip1.SelectedItem.Index - 1).Visible = True
        Frame2(iFrame2Atual).Visible = False
        iFrame2Atual = TabStrip1.SelectedItem.Index - 1
    End If
    
End Sub

Private Function Trata_Dados_Export_XLS(ByVal objBrowseExcel As AdmBrowseExcel) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objBrowseUsuCampo As AdmBrowseUsuarioCampo
Dim objBrowseExcelAux As AdmBrowseExcelAux

On Error GoTo Erro_Trata_Dados_Export_XLS

    FormFormulas.Clear
    FormFormulas.AddItem EXCEL_FORMULA_SUM_TEXTO
    FormFormulas.AddItem EXCEL_FORMULA_COUNT_TEXTO
    FormFormulas.AddItem EXCEL_FORMULA_MAX_TEXTO
    FormFormulas.AddItem EXCEL_FORMULA_MIN_TEXTO
    FormFormulas.AddItem EXCEL_FORMULA_AVG_TEXTO
    
    TabCamposForm.Clear
    TabCamposForm.AddItem EXCEL_FORMULA_SUM_TEXTO
    TabCamposForm.AddItem EXCEL_FORMULA_COUNT_TEXTO
    TabCamposForm.AddItem EXCEL_FORMULA_MAX_TEXTO
    TabCamposForm.AddItem EXCEL_FORMULA_MIN_TEXTO
    TabCamposForm.AddItem EXCEL_FORMULA_AVG_TEXTO
    
    TabCamposPosicao.Clear
    TabCamposPosicao.AddItem EXCEL_TABDIN_POS_FILTRO_TEXTO
    TabCamposPosicao.AddItem EXCEL_TABDIN_POS_LINHA_TEXTO
    TabCamposPosicao.AddItem EXCEL_TABDIN_POS_COLUNA_TEXTO
    TabCamposPosicao.AddItem EXCEL_TABDIN_POS_VALOR_TEXTO
    
    PlanTitulo.Text = objBrowseExcel.sTitulo
    
    If objBrowseExcel.iFormato = EXCEL_FORMATO_CSV Then
        FormatoCSV.Value = True
    Else
        FormatoXls.Value = True
    End If
    Call Trata_Formato
    
    If Len(Trim(objBrowseExcel.sArquivo)) = 0 Then
        Arquivo.Enabled = False
        NomeAuto.Value = vbChecked
    Else
        Arquivo.Enabled = True
        NomeAuto.Value = vbUnchecked
        Arquivo.Text = objBrowseExcel.sArquivo
    End If
    
    LocalizacaoCSV.Text = objBrowseExcel.sLocalizacaoCsv
    
    If objBrowseExcel.iTabelaDinamica = MARCADO Then
        IncluiTabela.Value = vbChecked
    Else
        IncluiTabela.Value = vbUnchecked
    End If
    Call IncluiTabela_Click
    
    If objBrowseExcel.iIncluirGrafico = MARCADO Then
        IncluirGrafico.Value = vbChecked
    Else
        IncluirGrafico.Value = vbUnchecked
    End If
    Call IncluirGrafico_Click

    TipoGrafico.Clear
    TipoGrafico.AddItem EXCEL_TIPOGRAFICO_AREA_TEXTO
    TipoGrafico.AddItem EXCEL_TIPOGRAFICO_COLUNA_TEXTO
    TipoGrafico.AddItem EXCEL_TIPOGRAFICO_LINHA_TEXTO
    TipoGrafico.AddItem EXCEL_TIPOGRAFICO_PIZZA_TEXTO

    Select Case objBrowseExcel.iTipoGrafico
        Case 0
            TipoGrafico.ListIndex = -1
        Case EXCEL_TIPOGRAFICO_AREA
            TipoGrafico.ListIndex = 0
        Case EXCEL_TIPOGRAFICO_COLUNA
            TipoGrafico.ListIndex = 1
        Case EXCEL_TIPOGRAFICO_LINHA
            TipoGrafico.ListIndex = 2
        Case EXCEL_TIPOGRAFICO_PIZZA
            TipoGrafico.ListIndex = 3
    End Select
    
    FormCampos.Clear
    TabCamposCampos.Clear
    
    iIndice = 0
    For Each objBrowseUsuCampo In colCamposDisponiveis
        FormCampos.AddItem objBrowseUsuCampo.sNome
        TabCamposCampos.AddItem objBrowseUsuCampo.sNome
        For Each objBrowseExcelAux In objBrowseExcel.colCampos
            If objBrowseExcelAux.sCampo = objBrowseUsuCampo.sNome Then
                iIndice = iIndice + 1
                GridTabCampos.TextMatrix(iIndice, COL_GRIDCAMPOS_CAMPO) = objBrowseUsuCampo.sNome
                Select Case objBrowseExcelAux.iFormula
                    Case 0
                        GridTabCampos.TextMatrix(iIndice, COL_GRIDCAMPOS_FORMULA) = ""
                    Case EXCEL_FORMULA_SUM
                        GridTabCampos.TextMatrix(iIndice, COL_GRIDCAMPOS_FORMULA) = EXCEL_FORMULA_SUM_TEXTO
                    Case EXCEL_FORMULA_COUNT
                        GridTabCampos.TextMatrix(iIndice, COL_GRIDCAMPOS_FORMULA) = EXCEL_FORMULA_COUNT_TEXTO
                    Case EXCEL_FORMULA_MIN
                        GridTabCampos.TextMatrix(iIndice, COL_GRIDCAMPOS_FORMULA) = EXCEL_FORMULA_MIN_TEXTO
                    Case EXCEL_FORMULA_MAX
                        GridTabCampos.TextMatrix(iIndice, COL_GRIDCAMPOS_FORMULA) = EXCEL_FORMULA_MAX_TEXTO
                    Case EXCEL_FORMULA_AVG
                        GridTabCampos.TextMatrix(iIndice, COL_GRIDCAMPOS_FORMULA) = EXCEL_FORMULA_AVG_TEXTO
                End Select
                Select Case objBrowseExcelAux.iPosicao
                    Case EXCEL_TABDIN_POS_FILTRO
                        GridTabCampos.TextMatrix(iIndice, COL_GRIDCAMPOS_POSICAO) = EXCEL_TABDIN_POS_FILTRO_TEXTO
                    Case EXCEL_TABDIN_POS_LINHA
                        GridTabCampos.TextMatrix(iIndice, COL_GRIDCAMPOS_POSICAO) = EXCEL_TABDIN_POS_LINHA_TEXTO
                    Case EXCEL_TABDIN_POS_VALOR
                        GridTabCampos.TextMatrix(iIndice, COL_GRIDCAMPOS_POSICAO) = EXCEL_TABDIN_POS_VALOR_TEXTO
                    Case EXCEL_TABDIN_POS_COLUNA
                        GridTabCampos.TextMatrix(iIndice, COL_GRIDCAMPOS_POSICAO) = EXCEL_TABDIN_POS_COLUNA_TEXTO
                End Select
            End If
        Next
    Next
    objGridCampos.iLinhasExistentes = iIndice
    
    iIndice = 0
    For Each objBrowseExcelAux In objBrowseExcel.colFormulas
        For Each objBrowseUsuCampo In colBrowseUsuarioCampo
            If objBrowseExcelAux.sCampo = objBrowseUsuCampo.sNome Then
                iIndice = iIndice + 1
                GridFormulas.TextMatrix(iIndice, COL_GRIDFORMULAS_CAMPO) = objBrowseExcelAux.sCampo
                Select Case objBrowseExcelAux.iFormula
                    Case EXCEL_FORMULA_SUM
                        GridFormulas.TextMatrix(iIndice, COL_GRIDFORMULAS_FORMULA) = EXCEL_FORMULA_SUM_TEXTO
                    Case EXCEL_FORMULA_COUNT
                        GridFormulas.TextMatrix(iIndice, COL_GRIDFORMULAS_FORMULA) = EXCEL_FORMULA_COUNT_TEXTO
                    Case EXCEL_FORMULA_MIN
                        GridFormulas.TextMatrix(iIndice, COL_GRIDFORMULAS_FORMULA) = EXCEL_FORMULA_MIN_TEXTO
                    Case EXCEL_FORMULA_MAX
                        GridFormulas.TextMatrix(iIndice, COL_GRIDFORMULAS_FORMULA) = EXCEL_FORMULA_MAX_TEXTO
                    Case EXCEL_FORMULA_AVG
                        GridFormulas.TextMatrix(iIndice, COL_GRIDFORMULAS_FORMULA) = EXCEL_FORMULA_AVG_TEXTO
                End Select
                Exit For
            End If
        Next
    Next
    objGridForm.iLinhasExistentes = iIndice
    
    Trata_Dados_Export_XLS = SUCESSO
    
    Exit Function
    
Erro_Trata_Dados_Export_XLS:

    Trata_Dados_Export_XLS = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143891)

    End Select

    Exit Function
    
End Function

Private Function Saida_Celula_Padrao(objGridInt As AdmGrid, ByVal objControle As Object, Optional ByVal bAdicionaLinha As Boolean = False, Optional ByVal bTestaRepeticao As Boolean = False) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinhasAnt As Integer
Dim iIndice As Integer
Dim iQtd As Integer

On Error GoTo Erro_Saida_Celula_Padrao

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.Text)) > 0 Then
    
        iLinhasAnt = objGridInt.iLinhasExistentes
        
        If bTestaRepeticao Then
            For iIndice = 1 To objGridInt.iLinhasExistentes
                If iIndice <> objGridInt.objGrid.Row Then
                    If UCase(objControle.Text) = UCase(objGridInt.objGrid.TextMatrix(iIndice, objGridInt.objGrid.Col)) Then gError 202145
                End If
            Next
        End If
       
        If bAdicionaLinha Then
            Call Adiciona_Linha(objGridInt)
        End If
        
        If objControle Is FormCampos And iLinhasAnt < objGridInt.iLinhasExistentes Then
            GridFormulas.TextMatrix(GridFormulas.Row, COL_GRIDFORMULAS_FORMULA) = EXCEL_FORMULA_SUM_TEXTO
        End If
        
        If objControle Is TabCamposCampos And iLinhasAnt < objGridInt.iLinhasExistentes Then
            GridTabCampos.TextMatrix(GridTabCampos.Row, COL_GRIDCAMPOS_POSICAO) = EXCEL_TABDIN_POS_FILTRO_TEXTO
        End If
        
        If objControle Is TabCamposPosicao Then
            If objControle.Text <> EXCEL_TABDIN_POS_VALOR_TEXTO Then
                GridTabCampos.TextMatrix(GridTabCampos.Row, COL_GRIDCAMPOS_FORMULA) = ""
            End If
            If Len(Trim(GridTabCampos.TextMatrix(GridTabCampos.Row, COL_GRIDCAMPOS_FORMULA))) = 0 Then
                If objControle.Text = EXCEL_TABDIN_POS_VALOR_TEXTO Then
                    GridTabCampos.TextMatrix(GridTabCampos.Row, COL_GRIDCAMPOS_FORMULA) = EXCEL_FORMULA_SUM_TEXTO
                End If
            End If
        End If
        
    End If
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 202074

    Saida_Celula_Padrao = SUCESSO

    Exit Function

Erro_Saida_Celula_Padrao:

    Saida_Celula_Padrao = gErr

    Select Case gErr

        Case 202074
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 202145
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_REPETIDO_NO_GRID", gErr, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202075)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Function Adiciona_Linha(ByVal objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Adiciona_Linha
              
    'verifica se precisa preencher o grid com uma nova linha
    If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
    
    Adiciona_Linha = SUCESSO
        
    Exit Function

Erro_Adiciona_Linha:

    Adiciona_Linha = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 202089)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable
              
    Select Case objControl.Name
                        
        Case TabCamposCampos.Name
            objControl.Enabled = True
            
        Case TabCamposPosicao.Name
            If Len(Trim(GridTabCampos.TextMatrix(GridTabCampos.Row, COL_GRIDCAMPOS_CAMPO))) <> 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
        Case TabCamposForm.Name
            If GridTabCampos.TextMatrix(GridTabCampos.Row, COL_GRIDCAMPOS_POSICAO) = EXCEL_TABDIN_POS_VALOR_TEXTO Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If

        Case FormCampos.Name
            objControl.Enabled = True
            
        Case FormFormulas.Name
            If Len(Trim(GridFormulas.TextMatrix(GridFormulas.Row, COL_GRIDFORMULAS_CAMPO))) <> 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If

    End Select
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 202141)

    End Select

    Exit Sub

End Sub

Public Sub Trata_Formato()

Dim lErro As Long

On Error GoTo Erro_Trata_Formato
        
    If FormatoXls.Value Then
        FrameForm.Enabled = True
        'LocalizacaoCSV.Text = ""
        'LocalizacaoCSV.Enabled = False
        IncluiTabela.Enabled = True
    Else
        FrameForm.Enabled = False
        Call Grid_Limpa(objGridForm)
        'LocalizacaoCSV.Enabled = True
        IncluiTabela.Value = vbUnchecked
        IncluiTabela.Enabled = False
    End If
    Call IncluiTabela_Click
        
    Exit Sub

Erro_Trata_Formato:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 202141)

    End Select

    Exit Sub

End Sub

Public Sub Trata_GridExcel(ByVal bRemove As Boolean, ByVal sCampo As String)

Dim lErro As Long
Dim iIndice As Integer
Dim iLinha As Integer
Dim objBrowseExcelAux As AdmBrowseExcelAux
Dim bAchou As Boolean

On Error GoTo Erro_Trata_GridExcel

    If bRemove Then
        For iIndice = 0 To FormCampos.ListCount - 1
            If sCampo = FormCampos.List(iIndice) Then
                FormCampos.RemoveItem (iIndice)
                Exit For
            End If
        Next
        iIndice = iIndice + 1
        For iLinha = iIndice To objGridCampos.iLinhasExistentes
            GridTabCampos.TextMatrix(iLinha, COL_GRIDCAMPOS_CAMPO) = GridTabCampos.TextMatrix(iLinha + 1, COL_GRIDCAMPOS_CAMPO)
            GridTabCampos.TextMatrix(iLinha, COL_GRIDCAMPOS_POSICAO) = GridTabCampos.TextMatrix(iLinha + 1, COL_GRIDCAMPOS_POSICAO)
            GridTabCampos.TextMatrix(iLinha, COL_GRIDCAMPOS_FORMULA) = GridTabCampos.TextMatrix(iLinha + 1, COL_GRIDCAMPOS_FORMULA)
        Next
        objGridCampos.iLinhasExistentes = objGridCampos.iLinhasExistentes - 1
        bAchou = False
        For iLinha = 1 To objGridForm.iLinhasExistentes
            If GridFormulas.TextMatrix(iLinha, COL_GRIDFORMULAS_CAMPO) = sCampo Then
                bAchou = True
            End If
            If bAchou Then
                GridFormulas.TextMatrix(iLinha, COL_GRIDFORMULAS_CAMPO) = GridFormulas.TextMatrix(iLinha + 1, COL_GRIDFORMULAS_CAMPO)
                GridFormulas.TextMatrix(iLinha, COL_GRIDFORMULAS_FORMULA) = GridFormulas.TextMatrix(iLinha + 1, COL_GRIDFORMULAS_FORMULA)
            End If
        Next
        objGridForm.iLinhasExistentes = objGridCampos.iLinhasExistentes - 1
    Else
        FormCampos.AddItem sCampo
        GridTabCampos.TextMatrix(objGridCampos.iLinhasExistentes + 1, COL_GRIDCAMPOS_CAMPO) = sCampo
        objGridCampos.iLinhasExistentes = objGridCampos.iLinhasExistentes + 1
    End If

    Exit Sub

Erro_Trata_GridExcel:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 202141)

    End Select

    Exit Sub

End Sub

Private Sub GridFormulas_Click()

Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridForm, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridForm, iAlterado)
    End If
    
End Sub

Private Sub GridFormulas_GotFocus()
    Call Grid_Recebe_Foco(objGridForm)
End Sub

Private Sub GridFormulas_EnterCell()
    Call Grid_Entrada_Celula(objGridForm, iAlterado)
End Sub

Private Sub GridFormulas_LeaveCell()
    Call Saida_Celula(objGridForm)
End Sub

Private Sub GridFormulas_KeyDown(KeyCode As Integer, Shift As Integer)

Dim dColunaSoma As Double
Dim lErro As Long

On Error GoTo Erro_GridLancamentos_KeyDown

    lErro = Grid_Trata_Tecla1(KeyCode, objGridForm)
    If lErro <> SUCESSO Then gError 44124
    
    Exit Sub
    
Erro_GridLancamentos_KeyDown:

    Select Case gErr
    
        Case 44124
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143898)
    
    End Select

    Exit Sub

End Sub

Private Sub GridFormulas_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridForm, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridForm, iAlterado)
    End If

End Sub

Private Sub GridFormulas_LostFocus()
    Call Grid_Libera_Foco(objGridForm)
End Sub

Private Sub GridFormulas_RowColChange()
    Call Grid_RowColChange(objGridForm)
End Sub

Private Sub GridFormulas_Scroll()
    Call Grid_Scroll(objGridForm)
End Sub

Private Sub GridTabCampos_Click()

Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridCampos, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCampos, iAlterado)
    End If
    
End Sub

Private Sub GridTabCampos_GotFocus()
    Call Grid_Recebe_Foco(objGridCampos)
End Sub

Private Sub GridTabCampos_EnterCell()
    Call Grid_Entrada_Celula(objGridCampos, iAlterado)
End Sub

Private Sub GridTabCampos_LeaveCell()
    Call Saida_Celula(objGridCampos)
End Sub

Private Sub GridTabCampos_KeyDown(KeyCode As Integer, Shift As Integer)

Dim dColunaSoma As Double
Dim lErro As Long

On Error GoTo Erro_GridLancamentos_KeyDown

    lErro = Grid_Trata_Tecla1(KeyCode, objGridCampos)
    If lErro <> SUCESSO Then gError 44124
    
    Exit Sub
    
Erro_GridLancamentos_KeyDown:

    Select Case gErr
    
        Case 44124
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143898)
    
    End Select

    Exit Sub

End Sub

Private Sub GridTabCampos_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCampos, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCampos, iAlterado)
    End If

End Sub

Private Sub GridTabCampos_LostFocus()
    Call Grid_Libera_Foco(objGridCampos)
End Sub

Private Sub GridTabCampos_RowColChange()
    Call Grid_RowColChange(objGridCampos)
End Sub

Private Sub GridTabCampos_Scroll()
    Call Grid_Scroll(objGridCampos)
End Sub

Private Sub BotaoProcurar_Click()

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

On Error GoTo Erro_BotaoProcurar_Click

    szTitle = "Localização física dos arquivos exportados"
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       
        LocalizacaoCSV.Text = sBuffer
        Call LocalizacaoCSV_Validate(bSGECancelDummy)
  
    End If
  
    Exit Sub

Erro_BotaoProcurar_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192326)

    End Select

    Exit Sub
  
End Sub

Private Sub LocalizacaoCSV_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iPos As Integer

On Error GoTo Erro_LocalizacaoCSV_Validate

    If Len(Trim(LocalizacaoCSV.Text)) = 0 Then Exit Sub
    
    If right(LocalizacaoCSV.Text, 1) <> "\" And right(LocalizacaoCSV.Text, 1) <> "/" Then
        iPos = InStr(1, LocalizacaoCSV.Text, "/")
        If iPos = 0 Then
            LocalizacaoCSV.Text = LocalizacaoCSV.Text & "\"
        Else
            LocalizacaoCSV.Text = LocalizacaoCSV.Text & "/"
        End If
    End If

    If Len(Trim(Dir(LocalizacaoCSV.Text, vbDirectory))) = 0 Then gError 192327

    Exit Sub

Erro_LocalizacaoCSV_Validate:

    Cancel = True

    Select Case gErr

        Case 192327, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, LocalizacaoCSV.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192328)

    End Select

    Exit Sub

End Sub

Private Sub NomeAuto_Click()
    If NomeAuto.Value = vbChecked Then
        Arquivo.Enabled = False
        Arquivo.Text = ""
    Else
        Arquivo.Enabled = True
    End If
End Sub
