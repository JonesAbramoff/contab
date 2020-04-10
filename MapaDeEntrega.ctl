VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl MapaDeEntregaOcx 
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   10995
   Begin VB.CheckBox OptGravarEntrega 
      Caption         =   "Gravar data de entrega no documento"
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
      Left            =   7170
      TabIndex        =   40
      Top             =   6900
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.CheckBox OptGravarTransp 
      Caption         =   "Gravar transportadora no documento"
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
      Left            =   3480
      TabIndex        =   39
      Top             =   6900
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.CommandButton BotaoMapa 
      Caption         =   "Exibir Mapa"
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
      Left            =   1755
      TabIndex        =   38
      Top             =   6810
      Width           =   1650
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10455
      Top             =   6330
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10455
      Top             =   5835
   End
   Begin VB.CommandButton BotaoVigensDoDia 
      Caption         =   "Viagens do dia"
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
      Left            =   30
      TabIndex        =   37
      ToolTipText     =   "Exibe todas as viagens do dia"
      Top             =   6825
      Width           =   1680
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transporte"
      Height          =   1515
      Index           =   2
      Left            =   45
      TabIndex        =   56
      Top             =   5265
      Width           =   10905
      Begin VB.ComboBox Transportadora 
         Height          =   315
         Left            =   960
         TabIndex        =   31
         Top             =   165
         Width           =   1965
      End
      Begin VB.Frame Frame4 
         Caption         =   "Última viagem do dia"
         Height          =   495
         Left            =   5850
         TabIndex        =   94
         Top             =   990
         Width           =   4950
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Retorno:"
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
            Index           =   8
            Left            =   2370
            TabIndex        =   101
            Top             =   225
            Width           =   870
         End
         Begin VB.Label VeiculoRetorno 
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   3315
            TabIndex        =   100
            Top             =   180
            Width           =   1005
         End
         Begin VB.Label VeiculoUltViag 
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1260
            TabIndex        =   96
            Top             =   180
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Index           =   6
            Left            =   315
            TabIndex        =   95
            Top             =   225
            Width           =   870
         End
      End
      Begin VB.CommandButton BotaoSugerir 
         Caption         =   "Sugerir Veículo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6315
         TabIndex        =   33
         ToolTipText     =   "Preenche automaticamente com um veículo que aguente a carga"
         Top             =   120
         Width           =   900
      End
      Begin VB.TextBox Responsavel 
         Height          =   315
         Left            =   8445
         MaxLength       =   100
         TabIndex        =   34
         Top             =   165
         Width           =   2370
      End
      Begin VB.Frame Frame1 
         Caption         =   "Disponibilidade"
         Height          =   495
         Index           =   5
         Left            =   5850
         TabIndex        =   75
         Top             =   510
         Width           =   4950
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Index           =   1
            Left            =   615
            TabIndex        =   79
            Top             =   225
            Width           =   585
         End
         Begin VB.Label VeiculoDispDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1260
            TabIndex        =   78
            Top             =   180
            Width           =   990
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Left            =   2415
            TabIndex        =   77
            Top             =   240
            Width           =   825
         End
         Begin VB.Label VeiculoDispAte 
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   3315
            TabIndex        =   76
            Top             =   180
            Width           =   990
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Horário"
         Height          =   495
         Index           =   4
         Left            =   105
         TabIndex        =   72
         Top             =   990
         Width           =   5685
         Begin MSMask.MaskEdBox HoraSaida 
            Height          =   300
            Left            =   870
            TabIndex        =   35
            Top             =   150
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox HoraRetorno 
            Height          =   300
            Left            =   3690
            TabIndex        =   36
            Top             =   150
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Saída:"
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
            Left            =   240
            TabIndex        =   74
            Top             =   195
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Retorno:"
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
            Index           =   5
            Left            =   2925
            TabIndex        =   73
            Top             =   195
            Width           =   750
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Capacidade"
         Height          =   495
         Index           =   3
         Left            =   105
         TabIndex        =   65
         Top             =   510
         Width           =   5685
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "m3"
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
            Left            =   4590
            TabIndex        =   71
            Top             =   210
            Width           =   240
         End
         Begin VB.Label VeiculoVolume 
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   3705
            TabIndex        =   70
            Top             =   165
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Volume:"
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
            Left            =   2865
            TabIndex        =   69
            Top             =   225
            Width           =   825
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "kg"
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
            Left            =   1755
            TabIndex        =   68
            Top             =   210
            Width           =   240
         End
         Begin VB.Label VeiculoPeso 
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   870
            TabIndex        =   67
            Top             =   165
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Peso:"
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
            Left            =   225
            TabIndex        =   66
            Top             =   210
            Width           =   585
         End
      End
      Begin MSMask.MaskEdBox Veiculo 
         Height          =   315
         Left            =   3810
         TabIndex        =   32
         Top             =   165
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label TransportadoraLabel 
         AutoSize        =   -1  'True
         Caption         =   "Transp.:"
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
         TabIndex        =   103
         Top             =   210
         Width           =   720
      End
      Begin VB.Label LabelResponsável 
         Alignment       =   1  'Right Justify
         Caption         =   "Responsável:"
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
         Left            =   7260
         TabIndex        =   82
         Top             =   210
         Width           =   1170
      End
      Begin VB.Label DescVeiculo 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4680
         TabIndex        =   64
         Top             =   165
         Width           =   1650
      End
      Begin VB.Label LabelVeiculo 
         Alignment       =   1  'Right Justify
         Caption         =   "Veículo:"
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
         Height          =   315
         Left            =   3015
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   63
         Top             =   195
         Width           =   780
      End
   End
   Begin VB.Frame FrameDoc 
      Caption         =   "Notas Fiscais"
      Height          =   4695
      Left            =   45
      TabIndex        =   47
      Top             =   585
      Width           =   10905
      Begin VB.Frame Frame2 
         Caption         =   "Ordenação"
         Height          =   3000
         Index           =   1
         Left            =   9720
         TabIndex        =   84
         Top             =   1620
         Visible         =   0   'False
         Width           =   1155
         Begin VB.CommandButton BotaoOrdAuto 
            Caption         =   "Auto"
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
            Left            =   90
            TabIndex        =   23
            Top             =   240
            Width           =   960
         End
         Begin VB.CommandButton BotaoFundo 
            Height          =   315
            Left            =   315
            Picture         =   "MapaDeEntrega.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1575
            Width           =   540
         End
         Begin VB.CommandButton BotaoSobe 
            Height          =   315
            Left            =   315
            Picture         =   "MapaDeEntrega.ctx":0312
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   900
            Width           =   540
         End
         Begin VB.CommandButton BotaoDesce 
            Height          =   315
            Left            =   315
            Picture         =   "MapaDeEntrega.ctx":04D4
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   1230
            Width           =   540
         End
         Begin VB.CommandButton BotaoTopo 
            Height          =   315
            Left            =   315
            Picture         =   "MapaDeEntrega.ctx":0696
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   570
            Width           =   540
         End
         Begin VB.Frame Frame3 
            Caption         =   "Troca"
            Height          =   1020
            Left            =   60
            TabIndex        =   85
            Top             =   1935
            Width           =   1035
            Begin VB.CommandButton BotaoMudaLinha 
               Caption         =   "Alterar"
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
               Left            =   30
               TabIndex        =   28
               Top             =   675
               Width           =   960
            End
            Begin MSMask.MaskEdBox LinhaDesejada 
               Height          =   315
               Left            =   540
               TabIndex        =   86
               Top             =   345
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   3
               Mask            =   "###"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "De"
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
               Index           =   10
               Left            =   45
               TabIndex        =   89
               Top             =   165
               Width           =   300
            End
            Begin VB.Label LinhaAtual 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   30
               TabIndex        =   88
               Top             =   345
               Width           =   480
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Para"
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
               Height          =   210
               Index           =   11
               Left            =   510
               TabIndex        =   87
               Top             =   165
               Width           =   450
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Seleção"
         Height          =   3000
         Index           =   0
         Left            =   9720
         TabIndex        =   83
         Top             =   1620
         Width           =   1155
         Begin VB.CommandButton BotaoCapacVeiculo 
            Caption         =   "Marcar confome o Veículo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   15
            TabIndex        =   22
            ToolTipText     =   "Marca de acordo com a capacidade do veículo"
            Top             =   2235
            Width           =   1095
         End
         Begin VB.CommandButton BotaoDesmarcarTodos 
            Caption         =   "Desmarcar Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   15
            Picture         =   "MapaDeEntrega.ctx":09A8
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Desmarca todas NFs"
            Top             =   1215
            Width           =   1095
         End
         Begin VB.CommandButton BotaoMarcarTodos 
            Caption         =   "Marcar Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   15
            Picture         =   "MapaDeEntrega.ctx":1B8A
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Marca todas as Nfs"
            Top             =   225
            Width           =   1095
         End
      End
      Begin MSMask.MaskEdBox NFFilial 
         Height          =   225
         Left            =   105
         TabIndex        =   97
         Top             =   1860
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin VB.CommandButton BotaoNF 
         Caption         =   "Nota Fiscal ..."
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
         Left            =   45
         TabIndex        =   29
         ToolTipText     =   "Abre a tela de edição da nota fiscal"
         Top             =   4275
         Width           =   1500
      End
      Begin VB.CommandButton BotaoTroca 
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
         Height          =   405
         Left            =   9705
         TabIndex        =   19
         ToolTipText     =   "Habilita as opções de ordenação"
         Top             =   1095
         Width           =   1155
      End
      Begin VB.CheckBox NFSel 
         DragMode        =   1  'Automatic
         Height          =   225
         Left            =   540
         TabIndex        =   62
         Top             =   1560
         Width           =   375
      End
      Begin VB.Frame FrameFiltro 
         Caption         =   "Filtros"
         Height          =   885
         Left            =   45
         TabIndex        =   57
         Top             =   165
         Width           =   10830
         Begin VB.CheckBox optSoPVsAbertos 
            Caption         =   "Trazer só pedidos abertos"
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
            Left            =   165
            TabIndex        =   8
            Top             =   645
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2640
         End
         Begin VB.Frame Frame1 
            Caption         =   "Data de Entrega Prevista"
            Height          =   645
            Index           =   1
            Left            =   6300
            TabIndex        =   104
            Top             =   105
            Width           =   3345
            Begin MSMask.MaskEdBox DataEntregaDe 
               Height          =   300
               Left            =   345
               TabIndex        =   13
               Top             =   255
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownEntregaDe 
               Height          =   300
               Left            =   1365
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   255
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEntregaAte 
               Height          =   300
               Left            =   2010
               TabIndex        =   15
               Top             =   255
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownEntregaAte 
               Height          =   300
               Left            =   3045
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   240
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.Label Label8 
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
               Left            =   1635
               TabIndex        =   106
               Top             =   315
               Width           =   360
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
               Left            =   30
               TabIndex        =   105
               Top             =   300
               Width           =   315
            End
         End
         Begin VB.CheckBox ManterNFs 
            Caption         =   "Manter Selecionados"
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
            Left            =   165
            TabIndex        =   7
            Top             =   435
            Value           =   1  'Checked
            Width           =   2235
         End
         Begin VB.CommandButton BotaoTrazerNF 
            Caption         =   "Trazer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   9660
            TabIndex        =   17
            Top             =   180
            Width           =   1155
         End
         Begin VB.ComboBox Regiao 
            Height          =   315
            Left            =   885
            TabIndex        =   6
            Top             =   135
            Width           =   1995
         End
         Begin VB.Frame Frame1 
            Caption         =   "Data de Emissão"
            Height          =   645
            Index           =   7
            Left            =   2895
            TabIndex        =   58
            Top             =   105
            Width           =   3390
            Begin MSMask.MaskEdBox DataDe 
               Height          =   300
               Left            =   345
               TabIndex        =   9
               Top             =   255
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDe 
               Height          =   300
               Left            =   1380
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   255
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataAte 
               Height          =   300
               Left            =   2040
               TabIndex        =   11
               Top             =   255
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownAte 
               Height          =   300
               Left            =   3090
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   240
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
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
               Height          =   195
               Left            =   30
               TabIndex        =   81
               Top             =   300
               Width           =   315
            End
            Begin VB.Label Label7 
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
               Left            =   1665
               TabIndex        =   80
               Top             =   315
               Width           =   360
            End
         End
         Begin VB.Label LabelRegiao 
            Alignment       =   1  'Right Justify
            Caption         =   "Região:"
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
            Left            =   30
            TabIndex        =   61
            Top             =   180
            Width           =   795
         End
      End
      Begin MSMask.MaskEdBox NFNum 
         Height          =   225
         Left            =   915
         TabIndex        =   48
         Top             =   1560
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NFCli 
         Height          =   225
         Left            =   1725
         TabIndex        =   49
         Top             =   1545
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NFReg 
         Height          =   225
         Left            =   3210
         TabIndex        =   50
         Top             =   1545
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NFBairro 
         Height          =   225
         Left            =   4320
         TabIndex        =   51
         Top             =   1530
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NFEnd 
         Height          =   225
         Left            =   5610
         TabIndex        =   52
         Top             =   1740
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NFKg 
         Height          =   225
         Left            =   7155
         TabIndex        =   53
         Top             =   1455
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NFVol 
         Height          =   225
         Left            =   7785
         TabIndex        =   54
         Top             =   1470
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridNF 
         Height          =   555
         Left            =   15
         TabIndex        =   18
         Top             =   1080
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   979
         _Version        =   393216
         Rows            =   21
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.CommandButton BotaoCliente 
         Caption         =   "Clientes ..."
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
         Left            =   1605
         TabIndex        =   30
         ToolTipText     =   "Abre a tela de cadastro de clientes"
         Top             =   4275
         Width           =   1500
      End
      Begin VB.Label QtdNFTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   8655
         TabIndex        =   99
         Top             =   4350
         Width           =   990
      End
      Begin VB.Label LabelQtd 
         Alignment       =   1  'Right Justify
         Caption         =   "Qtde NFs:"
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
         Height          =   225
         Left            =   7710
         TabIndex        =   98
         Top             =   4380
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Volume Total:"
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
         Height          =   225
         Index           =   2
         Left            =   5430
         TabIndex        =   93
         Top             =   4380
         Width           =   1200
      End
      Begin VB.Label VolumeTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   6645
         TabIndex        =   92
         Top             =   4350
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Peso Total:"
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
         Height          =   225
         Index           =   3
         Left            =   3150
         TabIndex        =   91
         Top             =   4380
         Width           =   1050
      End
      Begin VB.Label PesoTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   4215
         TabIndex        =   90
         Top             =   4350
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação"
      Height          =   600
      Index           =   0
      Left            =   45
      TabIndex        =   55
      Top             =   0
      Width           =   8205
      Begin VB.CheckBox OptImprime 
         Caption         =   "Imprimir ao gravar"
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
         Left            =   6285
         TabIndex        =   5
         Top             =   240
         Width           =   1845
      End
      Begin VB.ComboBox TipoDoc 
         Height          =   315
         ItemData        =   "MapaDeEntrega.ctx":2BA4
         Left            =   4575
         List            =   "MapaDeEntrega.ctx":2BAE
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   195
         Width           =   1695
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   1800
         Picture         =   "MapaDeEntrega.ctx":2BD0
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Numeração Automática"
         Top             =   210
         Width           =   300
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   930
         TabIndex        =   0
         Top             =   210
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   6
         Mask            =   "999999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Left            =   2730
         TabIndex        =   2
         Top             =   195
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Left            =   3840
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   195
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Doc:"
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
         Height          =   315
         Left            =   3750
         TabIndex        =   102
         Top             =   240
         Width           =   795
      End
      Begin VB.Label LabelData 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   1920
         TabIndex        =   60
         Top             =   255
         Width           =   780
      End
      Begin VB.Label LabelCodigo 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   105
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   59
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   8280
      ScaleHeight     =   450
      ScaleWidth      =   2595
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   75
      Width           =   2655
      Begin VB.CommandButton BotaoImprimir 
         Height          =   345
         Left            =   90
         Picture         =   "MapaDeEntrega.ctx":2CBA
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Imprimir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   585
         Picture         =   "MapaDeEntrega.ctx":2DBC
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1095
         Picture         =   "MapaDeEntrega.ctx":2F16
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1590
         Picture         =   "MapaDeEntrega.ctx":30A0
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2070
         Picture         =   "MapaDeEntrega.ctx":35D2
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
End
Attribute VB_Name = "MapaDeEntregaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim giContadorTempo As Integer
Dim giTentativa As Integer
Dim glNumIntRel As Long
Dim gsDiretorio As String

Dim gobjMapa As ClassMapaDeEntrega
Dim bDesabilitaCmdGridAux As Boolean
Dim bTrazendoDados As Boolean

Dim iTipoDocAnt As Integer

Dim objGridNF As AdmGrid
Dim iGrid_NFSel_Col As Integer
Dim iGrid_NFNum_Col As Integer
Dim iGrid_NFCli_Col As Integer
Dim iGrid_NFFilial_Col As Integer
Dim iGrid_NFReg_Col As Integer
Dim iGrid_NFBairro_Col As Integer
Dim iGrid_NFEnd_Col As Integer
Dim iGrid_NFKg_Col As Integer
Dim iGrid_NFVol_Col As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoVeiculo As AdmEvento
Attribute objEventoVeiculo.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Viagem de Entrega"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "MapaDeEntrega"

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

Private Sub BotaoMapa_Click()

Dim lErro As Long
Dim lNumIntRel As Long
Dim sDiretorio As String
Dim lRetorno As Long

On Error GoTo Erro_BotaoMapa_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If gobjFAT.iPossuiIntMapLink = DESMARCADO Then gError 205564
    
    lErro = Move_GridNF_Memoria(gobjMapa)
    If lErro <> SUCESSO Then gError 205565
    
    lErro = CF("Entrega_Exibe_Mapa_Prepara", gobjMapa, lNumIntRel)
    If lErro <> SUCESSO Then gError 205565

    sDiretorio = String(255, 0)
    lRetorno = GetPrivateProfileString("Forprint", "DirBin", "c:\sge\programa\", sDiretorio, 255, NOME_ARQUIVO_ADM)
    sDiretorio = left(sDiretorio, lRetorno)
    
    glNumIntRel = lNumIntRel
    gsDiretorio = sDiretorio

    lErro = WinExec(sDiretorio & "rota.exe 1 " & CStr(glEmpresa) & " " & CStr(lNumIntRel) & " 0 " & "Viagem_" & CStr(Codigo.Text), SW_NORMAL)

    Timer2.Enabled = True

    Exit Sub

Erro_BotaoMapa_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 205564
            Call Rotina_Aviso(vbOKOnly, "AVISO_FUNC_TERCEITOS_SEM_CONFIG")

        Case 205565

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205566)

    End Select
    
    Exit Sub
    
End Sub

Private Sub NFSel_Click()
    Call Calcula_Valores
End Sub

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
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objGridNF = Nothing
    Set gobjMapa = Nothing
    Set objEventoVeiculo = Nothing
    Set objEventoCodigo = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205262)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento
    Set objEventoVeiculo = New AdmEvento
    
    Set gobjMapa = New ClassMapaDeEntrega

    bDesabilitaCmdGridAux = False
    bTrazendoDados = False
    
    glNumIntRel = 0

    lErro = Inicializa_GridNF(objGridNF)
    If lErro <> SUCESSO Then gError 205263
    
    'Preenche Combo Regiao
    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê cada codigo e descricao da tabela RegioesVendas
    lErro = CF("Cod_Nomes_Le", "RegioesVendas", "Codigo", "Descricao", STRING_REGIAO_VENDA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 205321

    'preenche a ComboBox Regiao com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        Regiao.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Regiao.ItemData(Regiao.NewIndex) = objCodigoDescricao.iCodigo
    Next
    
    lErro = Carrega_Transportadoras
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    ManterNFs.Value = vbChecked
    
    TipoDoc.ListIndex = 0
    iTipoDocAnt = 0
    
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 205263, 205321
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205264)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objMapaDeEntrega As ClassMapaDeEntrega) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    bDesabilitaCmdGridAux = False

    If Not (objMapaDeEntrega Is Nothing) Then

        lErro = Traz_MapaDeEntrega_Tela(objMapaDeEntrega)
        If lErro <> SUCESSO Then gError 205265

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 205265

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205266)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objMapaDeEntrega As ClassMapaDeEntrega) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objMapaDeEntrega.iFilialEmpresa = giFilialEmpresa
    objMapaDeEntrega.lCodigo = StrParaLong(Codigo.Text)
    objMapaDeEntrega.dtData = StrParaDate(Data.Text)
    objMapaDeEntrega.iRegiao = Codigo_Extrai(Regiao.Text)
    objMapaDeEntrega.lVeiculo = StrParaLong(Veiculo.Text)
    objMapaDeEntrega.dVolumeTotal = StrParaDbl(VolumeTotal.Caption)
    objMapaDeEntrega.dPesoTotal = StrParaDbl(PesoTotal.Caption)
    If Len(Trim(HoraSaida.ClipText)) > 0 Then objMapaDeEntrega.dHoraSaida = CDate(HoraSaida.Text)
    If Len(Trim(HoraRetorno.ClipText)) > 0 Then objMapaDeEntrega.dHoraRetorno = CDate(HoraRetorno.Text)
    objMapaDeEntrega.sResponsavel = Responsavel.Text
    objMapaDeEntrega.iTransportadora = Codigo_Extrai(Transportadora.Text)
    objMapaDeEntrega.iTipoDoc = TipoDoc.ItemData(TipoDoc.ListIndex)
    
    lErro = Move_GridNF_Memoria(objMapaDeEntrega)
    If lErro <> SUCESSO Then gError 205267

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 205267

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205268)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objMapaDeEntrega As New ClassMapaDeEntrega

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "MapaDeEntrega"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objMapaDeEntrega)
    If lErro <> SUCESSO Then gError 205269

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colCampoValor.Add "Codigo", objMapaDeEntrega.lCodigo, 0, "Codigo"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 205269

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205270)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objMapaDeEntrega As New ClassMapaDeEntrega

On Error GoTo Erro_Tela_Preenche

    objMapaDeEntrega.iFilialEmpresa = giFilialEmpresa
    objMapaDeEntrega.lCodigo = colCampoValor.Item("Codigo").vValor

    If objMapaDeEntrega.iFilialEmpresa <> 0 And objMapaDeEntrega.lCodigo <> 0 Then

        lErro = Traz_MapaDeEntrega_Tela(objMapaDeEntrega)
        If lErro <> SUCESSO Then gError 205271

    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 205271

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205272)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objMapaDeEntrega As New ClassMapaDeEntrega
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Text)) = 0 Then gError 205273
    If Len(Trim(Veiculo.Text)) = 0 And Len(Trim(Transportadora.Text)) = 0 Then gError 205279

    'Preenche o objMapaDeEntrega
    lErro = Move_Tela_Memoria(objMapaDeEntrega)
    If lErro <> SUCESSO Then gError 205274
    
    If objMapaDeEntrega.colMapaDoc.Count = 0 Then gError 205278
    
    If StrParaDbl(VeiculoPeso.Caption) < StrParaDbl(PesoTotal.Caption) Or StrParaDbl(VeiculoVolume.Caption) < StrParaDbl(VolumeTotal.Caption) Then
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_VEICULO_ACIMA_DA_CAPACIDADE")
        If vbResult = vbNo Then gError 205276
    End If

    lErro = Trata_Alteracao(objMapaDeEntrega, objMapaDeEntrega.iFilialEmpresa, objMapaDeEntrega.lCodigo)
    If lErro <> SUCESSO Then gError 205275

    'Grava o/a MapaDeEntrega no Banco de Dados
    lErro = CF("MapaDeEntrega_Grava", objMapaDeEntrega)
    If lErro <> SUCESSO Then gError 205276
    
    If OptImprime.Value = vbChecked Then Call Viagem_Imprime(objMapaDeEntrega.lCodigo)

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 205273
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 205274 To 205276
        
        Case 205278 'SEM ITENS
             Call Rotina_Erro(vbOKOnly, "ERRO_VIAGEMENTREGA_SEM_NF", gErr)
       
        Case 205279 'Veiculo não preenchido
            Call Rotina_Erro(vbOKOnly, "ERRO_VEICULO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205277)

    End Select

    Exit Function

End Function

Function Limpa_Tela_MapaDeEntrega() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_MapaDeEntrega
   
    Regiao.ListIndex = -1
    DescVeiculo.Caption = ""
    VeiculoPeso.Caption = ""
    VeiculoVolume.Caption = ""
    VeiculoDispDe.Caption = ""
    VeiculoDispAte.Caption = ""
    VeiculoUltViag.Caption = ""
    VeiculoRetorno.Caption = ""
    PesoTotal.Caption = ""
    VolumeTotal.Caption = ""
    QtdNFTotal.Caption = ""
    
    Set gobjMapa = New ClassMapaDeEntrega
    glNumIntRel = 0

    Frame2(1).Visible = False
    Frame2(0).Visible = True
    BotaoTroca.Caption = "Ordenação"
    FrameFiltro.Enabled = True
            
    ManterNFs.Value = vbChecked

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    Call Grid_Limpa(objGridNF)
    
    TipoDoc.ListIndex = 0
    Call Trata_TipoDoc
    
    Transportadora.ListIndex = -1
    
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    iAlterado = 0

    Limpa_Tela_MapaDeEntrega = SUCESSO

    Exit Function

Erro_Limpa_Tela_MapaDeEntrega:

    Limpa_Tela_MapaDeEntrega = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205279)

    End Select

    Exit Function

End Function

Function Traz_MapaDeEntrega_Tela(ByVal objMapaDeEntrega As ClassMapaDeEntrega) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_MapaDeEntrega_Tela

    Call Limpa_Tela_MapaDeEntrega

    If objMapaDeEntrega.lCodigo <> 0 Then
        Codigo.PromptInclude = False
        Codigo.Text = CStr(objMapaDeEntrega.lCodigo)
        Codigo.PromptInclude = True
    End If

    'Lê o MapaDeEntrega que está sendo Passado
    lErro = CF("MapaDeEntrega_Le", objMapaDeEntrega, True)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 205280

    If lErro = SUCESSO Then

        If objMapaDeEntrega.lCodigo <> 0 Then
            Codigo.PromptInclude = False
            Codigo.Text = CStr(objMapaDeEntrega.lCodigo)
            Codigo.PromptInclude = True
        End If

        If objMapaDeEntrega.dtData <> DATA_NULA Then
            Data.PromptInclude = False
            Data.Text = Format(objMapaDeEntrega.dtData, "dd/mm/yy")
            Data.PromptInclude = True
        End If

        If objMapaDeEntrega.lVeiculo <> 0 Then
            Veiculo.PromptInclude = False
            Veiculo.Text = CStr(objMapaDeEntrega.lVeiculo)
            Veiculo.PromptInclude = True
            Call Veiculo_Validate(bSGECancelDummy)
        End If

        If objMapaDeEntrega.dVolumeTotal <> 0 Then
            VolumeTotal.Caption = Formata_Estoque(objMapaDeEntrega.dVolumeTotal)
        End If
        
        If objMapaDeEntrega.dPesoTotal <> 0 Then
            PesoTotal.Caption = Formata_Estoque(objMapaDeEntrega.dPesoTotal)
        End If

        If objMapaDeEntrega.dHoraSaida <> 0 Then
            HoraSaida.PromptInclude = False
            HoraSaida.Text = Format(objMapaDeEntrega.dHoraSaida, HoraSaida.Format)
            HoraSaida.PromptInclude = True
        End If

        If objMapaDeEntrega.dHoraRetorno <> 0 Then
            HoraRetorno.PromptInclude = False
            HoraRetorno.Text = Format(objMapaDeEntrega.dHoraRetorno, HoraRetorno.Format)
            HoraRetorno.PromptInclude = True
        End If

        Responsavel.Text = objMapaDeEntrega.sResponsavel
        
        Call Combo_Seleciona_ItemData(TipoDoc, objMapaDeEntrega.iTipoDoc)
        Call Combo_Seleciona_ItemData(Transportadora, objMapaDeEntrega.iTransportadora)
        
        lErro = Preenche_GridNF_Tela(objMapaDeEntrega)
        If lErro <> SUCESSO Then gError 205320
        
        Set gobjMapa = objMapaDeEntrega

        lErro = Move_GridNF_Memoria(gobjMapa)
        If lErro <> SUCESSO Then gError 205320

    End If

    iAlterado = 0

    Traz_MapaDeEntrega_Tela = SUCESSO

    Exit Function

Erro_Traz_MapaDeEntrega_Tela:

    Traz_MapaDeEntrega_Tela = gErr

    Select Case gErr

        Case 205280, 205320

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205281)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 205282

    'Limpa Tela
    Call Limpa_Tela_MapaDeEntrega

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 205282

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205283)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205284)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 205285

    Call Limpa_Tela_MapaDeEntrega

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 205285

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205286)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objMapaDeEntrega As New ClassMapaDeEntrega
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Text)) = 0 Then gError 205287

    objMapaDeEntrega.iFilialEmpresa = giFilialEmpresa
    objMapaDeEntrega.lCodigo = StrParaLong(Codigo.Text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_MAPADEENTREGA", objMapaDeEntrega.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("MapaDeEntrega_Exclui", objMapaDeEntrega)
        If lErro <> SUCESSO Then gError 205288

        'Limpa Tela
        Call Limpa_Tela_MapaDeEntrega

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 205287
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 205288

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205289)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se Codigo está preenchida
    If Len(Trim(Codigo.Text)) <> 0 Then

       'Critica a Codigo
       lErro = Long_Critica(Codigo.Text)
       If lErro <> SUCESSO Then gError 205290

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 205290

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205291)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 205292

        Data.Text = sData
        
        Call Veiculo_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 205292

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205293)

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
        If lErro <> SUCESSO Then gError 205294

        Data.Text = sData

        Call Veiculo_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 205294

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205295)

    End Select

    Exit Sub

End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)
    
End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data.ClipText)) <> 0 Then

        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError 205296

        Call Veiculo_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 205296

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205297)

    End Select

    Exit Sub

End Sub

Private Sub Data_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Regiao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Regiao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Regiao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim iCodigo As Integer

On Error GoTo Erro_Regiao_Validate

    'Verifica se foi preenchido o campo Regiao
    If Len(Trim(Regiao.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Regiao
    If Regiao.Text = Regiao.List(Regiao.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Regiao, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 205325

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objRegiaoVenda.iCodigo = iCodigo

        'Tenta ler Regiao de Venda com esse código no BD
        lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
        If lErro <> SUCESSO And lErro <> 16137 Then gError 205326
        
        'Não encontrou Regiao Venda BD
        If lErro <> SUCESSO Then gError 205327
        
        'Encontrou Regiao Venda no BD, coloca no Text da Combo
        Regiao.Text = CStr(objRegiaoVenda.iCodigo) & SEPARADOR & objRegiaoVenda.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 205328

    Exit Sub

Erro_Regiao_Validate:
    
    Cancel = True
    
    Select Case gErr

    Case 205325, 205326

    Case 205327  'Não encontrou RegiaoVenda no BD
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_REGIAO")

        If vbMsgRes = vbYes Then
            'Chama a tela RegiaoVenda
            Call Chama_Tela("RegiaoVenda", objRegiaoVenda)
        End If

    Case 205328
        Call Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_ENCONTRADA", gErr, Regiao.Text)

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205329)

    End Select

    Exit Sub

End Sub

Private Sub Veiculo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVeiculo As New ClassVeiculos
Dim lUltViagem As Long
Dim dHoraRetorno As Double

On Error GoTo Erro_Veiculo_Validate

    'Verifica se Veiculo está preenchida
    If Len(Trim(Veiculo.Text)) <> 0 Then

        objVeiculo.lCodigo = StrParaLong(Veiculo.Text)
       
        lErro = CF("Veiculos_le", objVeiculo)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 205300
       
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 205357
       
        DescVeiculo.Caption = objVeiculo.sDescricao
        VeiculoPeso.Caption = Formata_Estoque(objVeiculo.dCapacidadeKg)
        VeiculoVolume.Caption = Formata_Estoque(objVeiculo.dVolumeM3)
        
        VeiculoDispDe.Caption = Format(objVeiculo.dDispPadraoDe, "hh:mm:ss")
        VeiculoDispAte.Caption = Format(objVeiculo.dDispPadraoAte, "hh:mm:ss")
       
        'Tem que obter os dados das viagens do dia
        
        lErro = CF("Veiculos_le_UltViagem", objVeiculo.lCodigo, StrParaDate(Data.Text), lUltViagem, dHoraRetorno)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 205300
        
        If lUltViagem <> 0 Then
            VeiculoUltViag.Caption = CStr(lUltViagem)
        Else
            VeiculoUltViag.Caption = ""
        End If
        If dHoraRetorno > 0 Then
            VeiculoRetorno.Caption = Format(dHoraRetorno, "hh:mm:ss")
        Else
            VeiculoRetorno.Caption = ""
        End If
       
    Else
        DescVeiculo.Caption = ""
        VeiculoPeso.Caption = ""
        VeiculoVolume.Caption = ""
        VeiculoDispDe.Caption = ""
        VeiculoDispAte.Caption = ""
        VeiculoUltViag.Caption = ""
        VeiculoRetorno.Caption = ""
    End If

    Exit Sub

Erro_Veiculo_Validate:

    Cancel = True

    Select Case gErr

        Case 205300
        
        Case 205357
            Call Rotina_Erro(vbOKOnly, "ERRO_VEICULO_NAO_CADASTRADO", gErr, objVeiculo.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205301)

    End Select

    Exit Sub

End Sub

Private Sub Veiculo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Veiculo, iAlterado)
    
End Sub

Private Sub Veiculo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub HoraSaida_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_HoraSaida_Validate

    'Verifica se HoraSaida está preenchida
    If Len(Trim(HoraSaida.ClipText)) <> 0 Then

       'Critica a HoraSaida
       lErro = Hora_Critica(HoraSaida.Text)
       If lErro <> SUCESSO Then gError 205306

    End If

    Exit Sub

Erro_HoraSaida_Validate:

    Cancel = True

    Select Case gErr

        Case 205306

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205307)

    End Select

    Exit Sub

End Sub

Private Sub HoraSaida_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(HoraSaida, iAlterado)
    
End Sub

Private Sub HoraSaida_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub HoraRetorno_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_HoraRetorno_Validate

    'Verifica se HoraRetorno está preenchida
    If Len(Trim(HoraRetorno.ClipText)) <> 0 Then

       'Critica a HoraRetorno
       lErro = Hora_Critica(HoraRetorno.Text)
       If lErro <> SUCESSO Then gError 205308

    End If

    Exit Sub

Erro_HoraRetorno_Validate:

    Cancel = True

    Select Case gErr

        Case 205308

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205309)

    End Select

    Exit Sub

End Sub

Private Sub HoraRetorno_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(HoraRetorno, iAlterado)
    
End Sub

Private Sub HoraRetorno_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Responsavel_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Responsavel_Validate

    'Verifica se Responsável está preenchida
    If Len(Trim(Responsavel.Text)) <> 0 Then

       '#######################################
       'CRITICA Responsável
       '#######################################

    End If

    Exit Sub

Erro_Responsavel_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205310)

    End Select

    Exit Sub

End Sub

Private Sub Responsavel_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMapaDeEntrega As ClassMapaDeEntrega

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objMapaDeEntrega = obj1

    'Mostra os dados do MapaDeEntrega na tela
    lErro = Traz_MapaDeEntrega_Tela(objMapaDeEntrega)
    If lErro <> SUCESSO Then gError 205311

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 205311

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205312)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objMapaDeEntrega As New ClassMapaDeEntrega
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objMapaDeEntrega.lCodigo = StrParaLong(Codigo.Text)

    End If

    Call Chama_Tela("MapaDeEntregaLista", colSelecao, objMapaDeEntrega, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205313)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_GridNF(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add (" ")
    objGrid.colColuna.Add ("Número")
    objGrid.colColuna.Add ("Cliente")
    objGrid.colColuna.Add ("Filial")
    objGrid.colColuna.Add ("Região")
    objGrid.colColuna.Add ("Bairro")
    objGrid.colColuna.Add ("Endereço")
    objGrid.colColuna.Add ("Kg")
    objGrid.colColuna.Add ("m3")

    'Controles que participam do Grid
    objGrid.colCampo.Add (NFSel.Name)
    objGrid.colCampo.Add (NFNum.Name)
    objGrid.colCampo.Add (NFCli.Name)
    objGrid.colCampo.Add (NFFilial.Name)
    objGrid.colCampo.Add (NFReg.Name)
    objGrid.colCampo.Add (NFBairro.Name)
    objGrid.colCampo.Add (NFEnd.Name)
    objGrid.colCampo.Add (NFKg.Name)
    objGrid.colCampo.Add (NFVol.Name)

    'Colunas do Grid
    iGrid_NFSel_Col = 1
    iGrid_NFNum_Col = 2
    iGrid_NFCli_Col = 3
    iGrid_NFFilial_Col = 4
    iGrid_NFReg_Col = 5
    iGrid_NFBairro_Col = 6
    iGrid_NFEnd_Col = 7
    iGrid_NFKg_Col = 8
    iGrid_NFVol_Col = 9

    objGrid.objGrid = GridNF

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 500 + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 12

    'Largura da primeira coluna
    GridNF.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridNF = SUCESSO

End Function

Private Sub GridNF_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridNF, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNF, iAlterado)
    End If

End Sub

Private Sub GridNF_GotFocus()
    Call Grid_Recebe_Foco(objGridNF)
End Sub

Private Sub GridNF_EnterCell()
    If Not bDesabilitaCmdGridAux Then
        Call Grid_Entrada_Celula(objGridNF, iAlterado)
    End If
End Sub

Private Sub GridNF_LeaveCell()
    If Not bDesabilitaCmdGridAux Then
        Call Saida_Celula(objGridNF)
    End If
End Sub

Private Sub GridNF_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridNF, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNF, iAlterado)
    End If

End Sub

Private Sub GridNF_RowColChange()
    If Not bDesabilitaCmdGridAux Then
        Call Grid_RowColChange(objGridNF)
    End If
    
    LinhaAtual.Caption = CStr(GridNF.Row)

End Sub

Private Sub GridNF_Scroll()
    Call Grid_Scroll(objGridNF)
End Sub

Private Sub GridNF_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridNF)
End Sub

Private Sub GridNF_LostFocus()
    Call Grid_Libera_Foco(objGridNF)
End Sub

Private Sub NFSel_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NFSel_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridNF)
End Sub

Private Sub NFSel_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNF)
End Sub

Private Sub NFSel_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNF.objControle = NFSel
    lErro = Grid_Campo_Libera_Foco(objGridNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub NFNum_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NFNum_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridNF)
End Sub

Private Sub NFNum_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNF)
End Sub

Private Sub NFNum_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNF.objControle = NFNum
    lErro = Grid_Campo_Libera_Foco(objGridNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub NFCli_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NFCli_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridNF)
End Sub

Private Sub NFCli_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNF)
End Sub

Private Sub NFCli_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNF.objControle = NFCli
    lErro = Grid_Campo_Libera_Foco(objGridNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub NFReg_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NFReg_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridNF)
End Sub

Private Sub NFReg_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNF)
End Sub

Private Sub NFReg_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNF.objControle = NFReg
    lErro = Grid_Campo_Libera_Foco(objGridNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub NFBairro_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NFBairro_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridNF)
End Sub

Private Sub NFBairro_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNF)
End Sub

Private Sub NFBairro_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNF.objControle = NFBairro
    lErro = Grid_Campo_Libera_Foco(objGridNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub NFEnd_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NFEnd_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridNF)
End Sub

Private Sub NFEnd_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNF)
End Sub

Private Sub NFEnd_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNF.objControle = NFEnd
    lErro = Grid_Campo_Libera_Foco(objGridNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub NFKg_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NFKg_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridNF)
End Sub

Private Sub NFKg_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNF)
End Sub

Private Sub NFKg_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNF.objControle = NFKg
    lErro = Grid_Campo_Libera_Foco(objGridNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub NFVol_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NFVol_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridNF)
End Sub

Private Sub NFVol_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNF)
End Sub

Private Sub NFVol_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNF.objControle = NFVol
    lErro = Grid_Campo_Libera_Foco(objGridNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'GridNF
        If objGridInt.objGrid.Name = GridNF.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_NFSel_Col

                    lErro = Saida_Celula_Padrao(objGridInt, NFSel)
                    If lErro <> SUCESSO Then gError 205314

            End Select
                    
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 205315

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 205314

        Case 205315
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205316)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa o controle da coluna em questão
    Select Case objControl.Name

        Case NFSel.Name
            If BotaoTroca.Caption = "Ordenação" Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If

        Case Else
            objControl.Enabled = False

    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205317)

    End Select

    Exit Sub

End Sub

Function Preenche_GridNF_Tela(ByVal objMapaDeEntrega As ClassMapaDeEntrega, Optional ByVal bSoMarcados As Boolean = False) As Long

Dim lErro As Long
Dim objDoc As Object
Dim objFilial As ClassFilialCliente
Dim objCli As ClassCliente
Dim objEnd As ClassEndereco
Dim objMapaDoc As ClassMapaDeEntregaDoc
Dim objReg As ClassRegiaoVenda
Dim iLinha As Integer
Dim bCadastrado As Boolean
Dim iTotal As Integer

On Error GoTo Erro_Preenche_GridNF_Tela

    Call Grid_Limpa(objGridNF)
    
    iTotal = 0
    For Each objDoc In objMapaDeEntrega.colDocs
        bCadastrado = False
        For Each objMapaDoc In objMapaDeEntrega.colMapaDoc
            If objMapaDoc.lNumIntDoc = objDoc.lNumIntDoc Then
                bCadastrado = True
                Exit For
            End If
        Next
        If Not bSoMarcados Or bCadastrado Then
            iTotal = iTotal + 1
        End If
    Next
    
    If iTotal >= objGridNF.objGrid.Rows Then
        Call Refaz_Grid(objGridNF, iTotal)
    End If

    iLinha = 0
    For Each objDoc In objMapaDeEntrega.colDocs
        
        Set objCli = New ClassCliente
        Set objEnd = New ClassEndereco
        Set objFilial = New ClassFilialCliente
        Set objReg = New ClassRegiaoVenda
        
        bCadastrado = False
        For Each objMapaDoc In objMapaDeEntrega.colMapaDoc
            If objMapaDoc.lNumIntDoc = objDoc.lNumIntDoc Then
                bCadastrado = True
                Exit For
            End If
        Next
        
        If Not bSoMarcados Or bCadastrado Then
            
            iLinha = iLinha + 1
            
            objCli.lCodigo = objDoc.lCliente
            lErro = CF("Cliente_Le", objCli)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 205330
            
            objFilial.lCodCliente = objDoc.lCliente
            If iTipoDocAnt = MAPAENTREGA_TIPODOC_PV Then
                objFilial.iCodFilial = objDoc.iFilial
            Else
                objFilial.iCodFilial = objDoc.iFilialCli
            End If
            
            lErro = CF("FilialCliente_Le", objFilial)
            If lErro <> SUCESSO And lErro <> 12567 Then gError 205331
            
            objEnd.lCodigo = objFilial.lEnderecoEntrega
            lErro = CF("Endereco_le", objEnd)
            If lErro <> SUCESSO Then gError 205332
            If Len(Trim(objEnd.sEndereco)) = 0 Then
                objEnd.lCodigo = objFilial.lEndereco
                lErro = CF("Endereco_le", objEnd)
                If lErro <> SUCESSO Then gError 205333
            End If
            
            If bCadastrado Then
                GridNF.TextMatrix(iLinha, iGrid_NFSel_Col) = CStr(MARCADO)
            Else
                GridNF.TextMatrix(iLinha, iGrid_NFSel_Col) = CStr(DESMARCADO)
            End If
            
            If iTipoDocAnt = MAPAENTREGA_TIPODOC_PV Then
                GridNF.TextMatrix(iLinha, iGrid_NFNum_Col) = CStr(objDoc.lCodigo)
            Else
                GridNF.TextMatrix(iLinha, iGrid_NFNum_Col) = CStr(objDoc.lNumNotaFiscal)
            End If
            
            GridNF.TextMatrix(iLinha, iGrid_NFCli_Col) = CStr(objCli.lCodigo) & SEPARADOR & objCli.sNomeReduzido
            GridNF.TextMatrix(iLinha, iGrid_NFFilial_Col) = CStr(objFilial.iCodFilial) & SEPARADOR & objFilial.sNome
            
            If objFilial.iRegiao <> 0 Then
                objReg.iCodigo = objFilial.iRegiao
                lErro = CF("RegiaoVenda_Le", objReg)
                If lErro <> SUCESSO And lErro <> 16137 Then gError 205333
                
                GridNF.TextMatrix(iLinha, iGrid_NFReg_Col) = CStr(objReg.iCodigo) & SEPARADOR & objReg.sDescricao
            Else
                GridNF.TextMatrix(iLinha, iGrid_NFReg_Col) = ""
            End If
            
            GridNF.TextMatrix(iLinha, iGrid_NFBairro_Col) = objEnd.sBairro
            GridNF.TextMatrix(iLinha, iGrid_NFEnd_Col) = objEnd.sEndereco
            GridNF.TextMatrix(iLinha, iGrid_NFKg_Col) = Formata_Estoque(objDoc.dPesoBruto)
            GridNF.TextMatrix(iLinha, iGrid_NFVol_Col) = Formata_Estoque(objDoc.dVolumeTotal)
            
        End If
    
    Next
    objGridNF.iLinhasExistentes = iLinha
    
    Call Grid_Refresh_Checkbox(objGridNF)
    
    Call Calcula_Valores

    Preenche_GridNF_Tela = SUCESSO

    Exit Function

Erro_Preenche_GridNF_Tela:

    Preenche_GridNF_Tela = gErr

    Select Case gErr
    
        Case 205330 To 205333

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205318)

    End Select

    Exit Function

End Function

Function Move_GridNF_Memoria(ByVal objMapaDeEntrega As ClassMapaDeEntrega) As Long

Dim lErro As Long
Dim objDoc As Object
Dim objMapaDoc As ClassMapaDeEntregaDoc
Dim lSeq As Long
Dim colDocs As New Collection
Dim colMapaDoc As New Collection
Dim iLinha As Integer

On Error GoTo Erro_Move_GridNF_Memoria

    lSeq = 0
    For iLinha = 1 To objGridNF.iLinhasExistentes
        
        If StrParaInt(GridNF.TextMatrix(iLinha, iGrid_NFSel_Col)) = MARCADO Then
            lSeq = lSeq + 1
            Set objDoc = gobjMapa.colDocs.Item(iLinha)
            objDoc.lVolumeQuant = lSeq
            Set objMapaDoc = New ClassMapaDeEntregaDoc
            objMapaDoc.lNumIntDoc = objDoc.lNumIntDoc
            objMapaDoc.lSeq = lSeq
            colDocs.Add objDoc
            colMapaDoc.Add objMapaDoc
        End If

    Next
    
    Set objMapaDeEntrega.colMapaDoc = colMapaDoc
    Set objMapaDeEntrega.colDocs = colDocs
    
    Move_GridNF_Memoria = SUCESSO

    Exit Function

Erro_Move_GridNF_Memoria:

    Move_GridNF_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205319)

    End Select

    Exit Function

End Function

Private Sub BotaoTrazerNF_Click()

Dim lErro As Long
Dim objMapa As New ClassMapaDeEntrega
Dim iLinha As Integer
Dim objDoc As Object
Dim objDocAux As Object
Dim objMapaDoc As ClassMapaDeEntregaDoc
Dim bAchou As Boolean, iStatus As Integer

On Error GoTo Erro_BotaoTrazerNF_Click
    
    If iTipoDocAnt = MAPAENTREGA_TIPODOC_PV Then
        If optSoPVsAbertos.Value = vbChecked Then iStatus = STATUS_ABERTO
        lErro = CF("MapaDeEntrega_Le_PVs", objMapa, Codigo_Extrai(Regiao.Text), StrParaDate(DataDe.Text), StrParaDate(DataAte.Text), StrParaDate(DataEntregaDe.Text), StrParaDate(DataEntregaAte.Text), iStatus)
    Else
        lErro = CF("MapaDeEntrega_Le_NFs", objMapa, Codigo_Extrai(Regiao.Text), StrParaDate(DataDe.Text), StrParaDate(DataAte.Text), StrParaDate(DataEntregaDe.Text), StrParaDate(DataEntregaAte.Text))
    End If
    If lErro <> SUCESSO Then gError 205323
    
    Set gobjMapa.colMapaDoc = New Collection
    If ManterNFs.Value = vbChecked Then
        'Remove as NFs não selecionadas
        For iLinha = objGridNF.iLinhasExistentes To 1 Step -1
            If GridNF.TextMatrix(iLinha, iGrid_NFSel_Col) = CStr(DESMARCADO) Then
                gobjMapa.colDocs.Remove (iLinha)
            End If
        Next
        'Inclui MapaNF para informar que está pré cadastradas (vão vir marcadas)
        For Each objDoc In gobjMapa.colDocs
            Set objMapaDoc = New ClassMapaDeEntregaDoc
            objMapaDoc.lNumIntDoc = objDoc.lNumIntDoc
            gobjMapa.colMapaDoc.Add objMapaDoc
        Next
        'Inclui as NFs que vieram na leitura nova desde que não estejam na tela
        For Each objDoc In objMapa.colDocs
            bAchou = False
            For Each objDocAux In gobjMapa.colDocs
                If objDoc.lNumIntDoc = objDocAux.lNumIntDoc Then
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then gobjMapa.colDocs.Add objDoc
        Next
    Else
        Set gobjMapa.colDocs = objMapa.colDocs
    End If
    
    'Traz os dados para tela
    lErro = Preenche_GridNF_Tela(gobjMapa)
    If lErro <> SUCESSO Then gError 205324

    Exit Sub

Erro_BotaoTrazerNF_Click:

    Select Case gErr
    
        Case 205323, 205324

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205322)

    End Select

    Exit Sub

End Sub

Private Sub BotaoDesmarcarTodos_Click()
    Call Grid_Marca_Desmarca(objGridNF, iGrid_NFSel_Col, DESMARCADO)
    QtdNFTotal.Caption = "0"
    PesoTotal.Caption = "0,00"
    VolumeTotal.Caption = "0,00"
End Sub

Private Sub BotaoMarcarTodos_Click()
    Call Grid_Marca_Desmarca(objGridNF, iGrid_NFSel_Col, MARCADO)
    Call Calcula_Valores
End Sub

Private Sub BotaoTroca_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoTroca_Click

    If BotaoTroca.Caption = "Ordenação" Then
        Frame2(1).Visible = True
        Frame2(0).Visible = False
        BotaoTroca.Caption = "Seleção"
        
        lErro = Move_GridNF_Memoria(gobjMapa)
        If lErro <> SUCESSO Then gError 205261
        
        lErro = Preenche_GridNF_Tela(gobjMapa)
        If lErro <> SUCESSO Then gError 205261
        
        FrameFiltro.Enabled = False
        
    Else
        Frame2(1).Visible = False
        Frame2(0).Visible = True
        BotaoTroca.Caption = "Ordenação"
        
        FrameFiltro.Enabled = True
        
    End If
    
    Exit Sub

Erro_BotaoTroca_Click:

    Select Case gErr

        Case 205261
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 205340)
    
    End Select

    Exit Sub
    
End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    lErro = CF("Config_ObterAutomatico", "FATConfig", "NUM_PROX_MAPAENTREGA", "MapaDeEntrega", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 205339

    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 205339
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 205340)
    
    End Select

    Exit Sub

End Sub

Private Sub Calcula_Valores()
    Call Soma_Coluna_Grid(objGridNF, iGrid_NFSel_Col, QtdNFTotal, False, iGrid_NFSel_Col)
    Call Soma_Coluna_Grid(objGridNF, iGrid_NFKg_Col, PesoTotal, False, iGrid_NFSel_Col)
    Call Soma_Coluna_Grid(objGridNF, iGrid_NFVol_Col, VolumeTotal, False, iGrid_NFSel_Col)
    QtdNFTotal.Caption = Format(QtdNFTotal.Caption, "#,##0")
End Sub

Private Sub BotaoCapacVeiculo_Click()

Dim lErro As Long
Dim dPeso As Double
Dim dVolume As Double
Dim dPesoAnt As Double
Dim dVolumeAnt As Double
Dim iLinha As Integer
Dim iLinhaP As Integer
Dim bPula As Boolean

On Error GoTo Erro_BotaoCapacVeiculo_Click

    If Len(Trim(Veiculo.Text)) = 0 Then gError 205341
    If StrParaDbl(VeiculoPeso.Caption) = 0 And StrParaDbl(VeiculoVolume.Caption) = 0 Then gError 205342

    bPula = False
    For iLinha = 1 To objGridNF.iLinhasExistentes
        dPeso = dPeso + StrParaDbl(GridNF.TextMatrix(iLinha, iGrid_NFKg_Col))
        dVolume = dVolume + StrParaDbl(GridNF.TextMatrix(iLinha, iGrid_NFVol_Col))
        
        If dPeso > StrParaDbl(VeiculoPeso.Caption) And StrParaDbl(VeiculoPeso.Caption) <> 0 Then bPula = True
        If dVolume > StrParaDbl(VeiculoVolume.Caption) And StrParaDbl(VeiculoVolume.Caption) <> 0 Then bPula = True
        
        If Not bPula Then
            iLinhaP = iLinhaP + 1
            dPesoAnt = dPeso
            dVolumeAnt = dVolume
            GridNF.TextMatrix(iLinha, iGrid_NFSel_Col) = CStr(MARCADO)
        Else
            GridNF.TextMatrix(iLinha, iGrid_NFSel_Col) = CStr(DESMARCADO)
        End If
    Next
    
    Call Grid_Refresh_Checkbox(objGridNF)

    QtdNFTotal.Caption = Format(iLinhaP, "#,##0")
    PesoTotal.Caption = Format(dPesoAnt, "STANDARD")
    VolumeTotal.Caption = Format(dVolumeAnt, "STANDARD")

    Exit Sub

Erro_BotaoCapacVeiculo_Click:

    Select Case gErr
    
        Case 205341 'Veículo não informado
            Call Rotina_Erro(vbOKOnly, "ERRO_VEICULO_NAO_PREENCHIDO", gErr)
    
        Case 205342 'Veículo sem capacidade informada
            Call Rotina_Erro(vbOKOnly, "ERRO_VEICULO_SEM_CAPACIDADE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 205340)
    
    End Select

    Exit Sub

End Sub

Private Sub DataAte_GotFocus()
     Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)
End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtDataDe As Date
Dim dtDataAte As Date

On Error GoTo Erro_DataAte_Validate

    'Verifica se a Data Final foi digitada
    If Len(Trim(DataAte.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 205343
    
    'Compara com a data Final
    If Len(Trim(DataDe.ClipText)) > 0 Then
        
        dtDataDe = CDate(DataDe.Text)
        dtDataAte = CDate(DataAte.Text)
        
        If dtDataDe > dtDataAte Then gError 205344

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 205343
        
        Case 205344
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205345)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_GotFocus()
     Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)
End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtDataDe As Date
Dim dtDataAte As Date

On Error GoTo Erro_DataDe_Validate

    'Verifica se a Data Inicial foi digitada
    If Len(Trim(DataDe.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 205346
    
    'Compara com a data Fianal
    If Len(Trim(DataAte.ClipText)) > 0 Then
        
        dtDataDe = CDate(DataDe.Text)
        dtDataAte = CDate(DataAte.Text)
        
        If dtDataDe > dtDataAte Then gError 205347

    End If


    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 205346
        
        Case 205347
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205348)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownAte_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro Then gError 205349

    Exit Sub

Erro_UpDownAte_DownClick:

    Select Case gErr

        Case 205349

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205350)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownAte_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro Then gError 205351

    Exit Sub

Erro_UpDownAte_UpClick:

    Select Case gErr

        Case 205351

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205352)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDe_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro Then gError 205353

    Exit Sub

Erro_UpDownDe_DownClick:

    Select Case gErr

        Case 205353

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205354)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDe_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro Then gError 205355

    Exit Sub

Erro_UpDownDe_UpClick:

    Select Case gErr

        Case 205355

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205356)

    End Select

    Exit Sub

End Sub

Private Sub LabelVeiculo_Click()

Dim lErro As Long
Dim objVeiculo As New ClassVeiculos
Dim colSelecao As New Collection

On Error GoTo Erro_LabelVeiculo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Veiculo.Text)) <> 0 Then

        objVeiculo.lCodigo = StrParaLong(Veiculo.Text)

    End If

    Call Chama_Tela("VeiculosLista", colSelecao, objVeiculo, objEventoVeiculo)

    Exit Sub

Erro_LabelVeiculo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205313)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoVeiculo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objVeiculo As ClassVeiculos

On Error GoTo Erro_objEventoVeiculo_evSelecao

    Set objVeiculo = obj1

    Veiculo.PromptInclude = False
    Veiculo.Text = CStr(objVeiculo.lCodigo)
    Veiculo.PromptInclude = True
    Call Veiculo_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoVeiculo_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205312)

    End Select

    Exit Sub

End Sub

Private Sub BotaoSugerir_Click()

Dim lErro As Long
Dim objVeiculo As New ClassVeiculos

On Error GoTo Erro_BotaoSugerir_Click

    lErro = CF("Veiculos_Le_Capacidade_Prox", objVeiculo, StrParaDbl(PesoTotal.Caption), StrParaDbl(VolumeTotal.Caption))
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 205359
    
    If lErro = ERRO_LEITURA_SEM_DADOS Then gError 205360

    Veiculo.PromptInclude = False
    Veiculo.Text = CStr(objVeiculo.lCodigo)
    Veiculo.PromptInclude = True
    Call Veiculo_Validate(bSGECancelDummy)

    Exit Sub

Erro_BotaoSugerir_Click:

    Select Case gErr
    
        Case 205359
        
        Case 205360
            Call Rotina_Erro(vbOKOnly, "ERRO_VEICULO_COM_CAPAC_NAO_ENCONTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205361)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoNF_Click()

Dim lErro As Long
Dim objNF As New ClassNFiscal
Dim objNFAux As ClassNFiscal
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iLinha As Integer
Dim objPV As New ClassPedidoDeVenda
Dim objPVAux As ClassPedidoDeVenda

On Error GoTo Erro_BotaoVerEtapas_Click
    
    'Se não tiver linha selecionada => Erro
    If GridNF.Row = 0 Then gError 205363
    
    If iTipoDocAnt = MAPAENTREGA_TIPODOC_PV Then
       
        Set objPVAux = gobjMapa.colDocs.Item(GridNF.Row)
    
        objPV.lCodigo = objPVAux.lCodigo
        objPV.iFilialEmpresa = objPVAux.iFilialEmpresa
    
        'Chama a tela de PV
        Call Chama_Tela("PedidoVenda", objPV)
    
    Else
    
        Set objNFAux = gobjMapa.colDocs.Item(GridNF.Row)
        objTipoDocInfo.iCodigo = objNFAux.iTipoNFiscal
       
        'lê o Tipo da Nota Fiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 205364
       
        objNF.lNumIntDoc = objNFAux.lNumIntDoc
        objNF.iFilialEmpresa = objNFAux.iFilialEmpresa
        objNF.lNumNotaFiscal = objNFAux.lNumNotaFiscal
        objNF.dtDataEmissao = objNFAux.dtDataEmissao
        objNF.sSerie = objNFAux.sSerie
    
        'Chama a tela de NF
        Call Chama_Tela(objTipoDocInfo.sNomeTelaNFiscal, objNF)
        
    End If

    Exit Sub

Erro_BotaoVerEtapas_Click:

    Select Case gErr

        Case 205363
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 205364

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205365)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoCliente_Click()

Dim lErro As Long
Dim objCli As New ClassCliente
Dim objDocAux As Object

On Error GoTo Erro_BotaoCliente_Click
    
    'Se não tiver linha selecionada => Erro
    If GridNF.Row = 0 Then gError 205363
       
    Set objDocAux = gobjMapa.colDocs.Item(GridNF.Row)
    
    objCli.lCodigo = objDocAux.lCliente
 
    'Chama a tela de ordem de produção
    Call Chama_Tela("Clientes", objCli)

    Exit Sub

Erro_BotaoCliente_Click:

    Select Case gErr

        Case 205366
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205367)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoDesce_Click()
    Call Troca_Dados_Posicao(GridNF.Row, GridNF.Row + 1)
End Sub

Private Sub BotaoFundo_Click()
    Call Troca_Dados_Posicao(GridNF.Row, objGridNF.iLinhasExistentes)
End Sub

Private Sub BotaoSobe_Click()
    Call Troca_Dados_Posicao(GridNF.Row, GridNF.Row - 1)
End Sub

Private Sub BotaoTopo_Click()
    Call Troca_Dados_Posicao(GridNF.Row, 1)
End Sub

Private Sub BotaoMudaLinha_Click()
    Call Troca_Dados_Posicao(GridNF.Row, StrParaInt(LinhaDesejada.Text))
End Sub

Private Function Troca_Dados_Posicao(ByVal iLinha1 As Integer, ByVal iLinha2 As Integer) As Long

Dim lErro As Long
Dim colOrd As New Collection, colCampos As New Collection
Dim objDoc As Object, iOrdem As Integer

On Error GoTo Erro_Troca_Dados_Posicao

    If iLinha1 < 1 Or iLinha1 > gobjMapa.colDocs.Count Then gError 205219
    If iLinha2 < 1 Or iLinha2 > gobjMapa.colDocs.Count Then gError 205220
  
    If iLinha1 > iLinha2 Then
        iOrdem = 0
        For Each objDoc In gobjMapa.colDocs
            iOrdem = iOrdem + 1
            If iOrdem < iLinha1 And iOrdem >= iLinha2 Then
                objDoc.lVolumeQuant = objDoc.lVolumeQuant + 1
            End If
            If iOrdem = iLinha1 Then objDoc.lVolumeQuant = iLinha2
        Next
    Else
        iOrdem = 0
        For Each objDoc In gobjMapa.colDocs
            iOrdem = iOrdem + 1
            If iOrdem > iLinha1 And iOrdem <= iLinha2 Then
                objDoc.lVolumeQuant = objDoc.lVolumeQuant - 1
            End If
            If iOrdem = iLinha1 Then objDoc.lVolumeQuant = iLinha2
        Next
    End If
    
    colCampos.Add "lVolumeQuant"
    Call Ordena_Colecao(gobjMapa.colDocs, colOrd, colCampos)
    
    Set gobjMapa.colDocs = colOrd
    
    lErro = Preenche_GridNF_Tela(gobjMapa)
    If lErro <> SUCESSO Then gError 205261
       
    bDesabilitaCmdGridAux = True
    GridNF.Row = iLinha2
    bDesabilitaCmdGridAux = False
    
    Troca_Dados_Posicao = SUCESSO
    
    Exit Function

Erro_Troca_Dados_Posicao:

    bDesabilitaCmdGridAux = False

    Troca_Dados_Posicao = gErr

    Select Case gErr
    
        Case 205219
        
        Case 205220
             Call Rotina_Erro(vbOKOnly, "ERRO_MUDANCA_LINHA_INVALIDA", gErr, iLinha2, 1, gobjMapa.colDocs.Count)
        
        Case 205261
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205218)

    End Select

    Exit Function

End Function

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is Veiculo Then
            Call LabelVeiculo_Click
        End If
    
    End If

End Sub

Private Sub LinhaDesejada_GotFocus()
Dim iAux As Integer
    Call MaskEdBox_TrataGotFocus(LinhaDesejada, iAux)
End Sub

Private Sub BotaoVigensDoDia_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim sFiltro As String
Dim objMapa As New ClassMapaDeEntrega

On Error GoTo Erro_BotaoVigensDoDia_Click

    colSelecao.Add StrParaDate(Data.Text)
    sFiltro = "Data = ?"

    Call Chama_Tela("MapaDeEntregaLista", colSelecao, objMapa, objEventoCodigo, sFiltro)

    Exit Sub

Erro_BotaoVigensDoDia_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205338)

    End Select

    Exit Sub

End Sub

Private Sub BotaoOrdAuto_Click()

Dim lErro As Long
Dim lNumIntRel As Long
Dim sDiretorio As String
Dim lRetorno As Long

On Error GoTo Erro_BotaoOrdAuto_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If gobjFAT.iPossuiIntMapLink = DESMARCADO Then gError 205564
    
    lErro = Move_GridNF_Memoria(gobjMapa)
    If lErro <> SUCESSO Then gError 205565
    
    lErro = CF("Entrega_Seq_Mapa_Prepara", gobjMapa, lNumIntRel)
    If lErro <> SUCESSO Then gError 205565

    sDiretorio = String(255, 0)
    lRetorno = GetPrivateProfileString("Forprint", "DirBin", "c:\sge\programa\", sDiretorio, 255, NOME_ARQUIVO_ADM)
    sDiretorio = left(sDiretorio, lRetorno)
    
    glNumIntRel = lNumIntRel
    gsDiretorio = sDiretorio

    lErro = WinExec(sDiretorio & "rota.exe 2 " & CStr(glEmpresa) & " " & CStr(lNumIntRel) & " 0 " & "Mapa_de_Entrega", SW_NORMAL)

    Timer1.Enabled = True

    Exit Sub

Erro_BotaoOrdAuto_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 205564
            Call Rotina_Aviso(vbOKOnly, "AVISO_FUNC_TERCEITOS_SEM_CONFIG")

        Case 205565

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205566)

    End Select
    
    Exit Sub
    
End Sub

Private Sub Timer1_Timer()

Const TEMPO_MAX_ESPERA = 40
Const NUM_TENTATIVAS = 2

Dim lErro As Long
Dim vbResult As VbMsgBoxResult
Dim sRetMsg As String
Dim lNumIntRel As Long
Dim lErro2 As Long

On Error GoTo Erro_Timer1_Timer

    giContadorTempo = giContadorTempo + 1
    
    If giContadorTempo > TEMPO_MAX_ESPERA Then
        GL_objMDIForm.MousePointer = vbDefault
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_TEMPO_ESPERA_ULTRAPASSADO", TEMPO_MAX_ESPERA)
        If vbResult = vbNo Then gError 205568
        GL_objMDIForm.MousePointer = vbHourglass
        giContadorTempo = 0
    End If
    
    lErro = CF("MapaRota1_Verifica_Retorno", glNumIntRel, sRetMsg)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then
    
        giTentativa = giTentativa + 1
        If giTentativa > NUM_TENTATIVAS Then gError 205567
    
        lErro2 = CF("Entrega_Seq_Mapa_Prepara", gobjMapa, lNumIntRel)
        If lErro2 <> SUCESSO Then gError 205568
        
        glNumIntRel = lNumIntRel
    
        Call WinExec(gsDiretorio & "rota.exe 2 " & CStr(glEmpresa) & " " & CStr(glNumIntRel) & " 1 " & "Rota_" & CStr(Codigo.Text), SW_NORMAL)
       
    End If
    
    If lErro = SUCESSO Then
        Timer1.Enabled = False
        giContadorTempo = 0
        
        lErro = CF("Entrega_Seq_Mapa_Obtem", gobjMapa, glNumIntRel)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 205569
        
        lErro = Preenche_GridNF_Tela(gobjMapa)
        If lErro <> SUCESSO Then gError 205570
        
        GL_objMDIForm.MousePointer = vbDefault
        
    End If

    Exit Sub

Erro_Timer1_Timer:

    Timer1.Enabled = False
    giContadorTempo = 0
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 205567
            Call Rotina_Erro(vbOKOnly, sRetMsg, gErr)
    
        Case 205568 To 205570

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205571)

    End Select
    
    Exit Sub
    
End Sub

Private Sub Timer2_Timer()

Const TEMPO_MAX_ESPERA = 40
Const NUM_TENTATIVAS = 2

Dim lErro As Long
Dim vbResult As VbMsgBoxResult
Dim sRetMsg As String
Dim lNumIntRel As Long
Dim lErro2 As Long

On Error GoTo Erro_Timer2_Timer

    giContadorTempo = giContadorTempo + 1
    
    If giContadorTempo > TEMPO_MAX_ESPERA Then
        GL_objMDIForm.MousePointer = vbDefault
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_TEMPO_ESPERA_ULTRAPASSADO", TEMPO_MAX_ESPERA)
        If vbResult = vbNo Then gError 205568
        GL_objMDIForm.MousePointer = vbHourglass
        giContadorTempo = 0
    End If
    
    lErro = CF("MapaRota1_Verifica_Retorno", glNumIntRel, sRetMsg)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then
    
        giTentativa = giTentativa + 1
        If giTentativa > NUM_TENTATIVAS Then gError 205567
    
        lErro2 = CF("Entrega_Exibe_Mapa_Prepara", gobjMapa, lNumIntRel)
        If lErro2 <> SUCESSO Then gError 205568
        
        glNumIntRel = lNumIntRel
    
        Call WinExec(gsDiretorio & "rota.exe 1 " & CStr(glEmpresa) & " " & CStr(glNumIntRel) & " 0 " & "Viagem_" & CStr(Codigo.Text), SW_NORMAL)
       
    End If
    
    If lErro = SUCESSO Then
        Timer2.Enabled = False
        giContadorTempo = 0
        
        GL_objMDIForm.MousePointer = vbDefault
        
    End If

    Exit Sub

Erro_Timer2_Timer:

    Timer2.Enabled = False
    giContadorTempo = 0
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 205567
            Call Rotina_Erro(vbOKOnly, sRetMsg, gErr)
    
        Case 205568 To 205570

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205571)

    End Select
    
    Exit Sub
    
End Sub

Private Function Trata_TipoDoc() As Long

    If TipoDoc.ListIndex <> -1 Then
        If iTipoDocAnt <> TipoDoc.ItemData(TipoDoc.ListIndex) Then
            
            Set gobjMapa = New ClassMapaDeEntrega
            Call Grid_Limpa(objGridNF)
            
            Frame2(1).Visible = False
            Frame2(0).Visible = True
            BotaoTroca.Caption = "Ordenação"
            FrameFiltro.Enabled = True
            
            PesoTotal.Caption = ""
            VolumeTotal.Caption = ""
            QtdNFTotal.Caption = ""
            
            iTipoDocAnt = TipoDoc.ItemData(TipoDoc.ListIndex)
            
            If iTipoDocAnt = MAPAENTREGA_TIPODOC_PV Then
                FrameDoc.Caption = "Pedidos de Venda"
                LabelQtd.Caption = "Qtde PVs:"
                BotaoNF.Caption = "Pedido ..."
                optSoPVsAbertos.Visible = True
                optSoPVsAbertos.Value = vbChecked
            Else
                FrameDoc.Caption = "Notas Fiscais"
                LabelQtd.Caption = "Qtde NFs:"
                BotaoNF.Caption = "Nota Fiscal ..."
                optSoPVsAbertos.Visible = False
                optSoPVsAbertos.Value = vbUnchecked
            End If
        End If
    End If

End Function

Private Sub TipoDoc_Change()
    Call Trata_TipoDoc
End Sub

Private Sub TipoDoc_Click()
    Call Trata_TipoDoc
End Sub

Private Sub DataEntregaAte_GotFocus()
     Call MaskEdBox_TrataGotFocus(DataEntregaAte, iAlterado)
End Sub

Private Sub DataEntregaAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtDataDe As Date
Dim dtDataEntregaAte As Date

On Error GoTo Erro_DataEntregaAte_Validate

    'Verifica se a Data Final foi digitada
    If Len(Trim(DataEntregaAte.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEntregaAte.Text)
    If lErro <> SUCESSO Then gError 205343
    
    'Compara com a data Final
    If Len(Trim(DataDe.ClipText)) > 0 Then
        
        dtDataDe = CDate(DataDe.Text)
        dtDataEntregaAte = CDate(DataEntregaAte.Text)
        
        If dtDataDe > dtDataEntregaAte Then gError 205344

    End If

    Exit Sub

Erro_DataEntregaAte_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 205343
        
        Case 205344
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205345)

    End Select

    Exit Sub

End Sub

Private Sub DataEntregaDe_GotFocus()
     Call MaskEdBox_TrataGotFocus(DataEntregaDe, iAlterado)
End Sub

Private Sub DataEntregaDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtDataDe As Date
Dim dtDataEntregaDe As Date

On Error GoTo Erro_DataEntregaDe_Validate

    'Verifica se a Data Final foi digitada
    If Len(Trim(DataEntregaDe.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEntregaDe.Text)
    If lErro <> SUCESSO Then gError 205343
    
    'Compara com a data Final
    If Len(Trim(DataDe.ClipText)) > 0 Then
        
        dtDataDe = CDate(DataDe.Text)
        dtDataEntregaDe = CDate(DataEntregaDe.Text)
        
        If dtDataDe > dtDataEntregaDe Then gError 205344

    End If

    Exit Sub

Erro_DataEntregaDe_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 205343
        
        Case 205344
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205345)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntregaAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntregaAte_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEntregaAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 205349

    Exit Sub

Erro_UpDownEntregaAte_DownClick:

    Select Case gErr

        Case 205349

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205350)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntregaAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntregaAte_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEntregaAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 205351

    Exit Sub

Erro_UpDownEntregaAte_UpClick:

    Select Case gErr

        Case 205351

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205352)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntregaDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntregaDe_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEntregaDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 205353

    Exit Sub

Erro_UpDownEntregaDe_DownClick:

    Select Case gErr

        Case 205353

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205354)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntregaDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntregaDe_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEntregaDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 205355

    Exit Sub

Erro_UpDownEntregaDe_UpClick:

    Select Case gErr

        Case 205355

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205356)

    End Select

    Exit Sub

End Sub

Private Function Carrega_Transportadoras() As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_Transportadoras

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_TRANSPORTADORA_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o item na Lista de Tabela de Preços
        Transportadora.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        Transportadora.ItemData(Transportadora.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_Transportadoras = SUCESSO

    Exit Function

Erro_Carrega_Transportadoras:

    Carrega_Transportadoras = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157916)

    End Select

    Exit Function

End Function

Public Sub Transportadora_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Transportadora_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Transportadora_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTransportadora As New ClassTransportadora
Dim iCodigo As Integer

On Error GoTo Erro_Transportadora_Validate

    'Verifica se foi preenchida a ComboBox Transportadora
    If Len(Trim(Transportadora.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox Transportadora
    If Transportadora.Text = Transportadora.List(Transportadora.ListIndex) Then Exit Sub

    'Verifica se existe o item na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Transportadora, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 26705

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 26706
        If lErro <> SUCESSO Then gError 26707 'Não encontrou Transportadora no BD

        'Encontrou Transportadora no BD, coloca no Text da Combo
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNome

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 26708

    Exit Sub

Erro_Transportadora_Validate:

    Cancel = True

    Select Case gErr

        Case 26705, 26706


        Case 26707  'Não encontrou Transportadora no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA", iCodigo)

            If vbMsgRes = vbYes Then

                Call Chama_Tela("Transportadora", objTransportadora)

            End If
            'Segura o foco

        Case 26708
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, Transportadora.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158002)

    End Select

    Exit Sub

End Sub

Public Sub BotaoImprimir_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoImprimir_Click

    'Se o código do orçamento não foi informado => erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 102238
    
    'Dispara função para imprimir orçamento
    lErro = Viagem_Imprime(Trim(Codigo.Text))
    If lErro <> SUCESSO Then gError 102239
    
    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 102239
        
        Case 102238
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 157836)

    End Select

    Exit Sub

End Sub

Private Function Viagem_Imprime(ByVal lCodigo As Long) As Long

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objMapa As New ClassMapaDeEntrega

On Error GoTo Erro_Viagem_Imprime

    'Transforma o ponteiro do mouse em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Guarda no obj o código do orçamento passado como parâmetro
    objMapa.lCodigo = lCodigo
    objMapa.iFilialEmpresa = giFilialEmpresa
    
    'Lê os dados do orçamento para verificar se o mesmo existe no BD
    lErro = CF("MapaDeEntrega_Le", objMapa)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM

    'Se não encontrou => erro, pois não é possível imprimir um orçamento inexistente
    If lErro = ERRO_LEITURA_SEM_DADOS Then gError 209068
    
    lErro = objRelatorio.ExecutarDireto("Romaneio de Entrega", "Codigo = @NCODIGO", 1, "RomEnt", "NCODIGO", CStr(lCodigo))
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault
    
    Viagem_Imprime = SUCESSO
    
    Exit Function

Erro_Viagem_Imprime:

    Viagem_Imprime = gErr
    
    Select Case gErr
           
        Case 209068
            Call Rotina_Erro(vbOKOnly, "ERRO_MAPADEENTREGA_NAO_CADASTRADO", gErr, objMapa.iFilialEmpresa, objMapa.lCodigo)
            
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209069)
    
    End Select
    
    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault

End Function

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    
    objGridInt.objGrid.Rows = iNumLinhas + 1
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
    
End Sub
