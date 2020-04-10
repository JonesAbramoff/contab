VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl PVConsulta2Ocx 
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   10995
   Begin VB.Frame Frame3 
      Caption         =   "Pedido de Venda Selecionado"
      Height          =   2565
      Left            =   75
      TabIndex        =   31
      Top             =   4230
      Width           =   10830
      Begin VB.CommandButton BotaoItens 
         Caption         =   "Itens"
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
         Left            =   1256
         TabIndex        =   18
         Top             =   2205
         Width           =   1050
      End
      Begin VB.CommandButton BotaoAnotacao 
         Caption         =   "Anotações"
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
         Left            =   7416
         TabIndex        =   23
         Top             =   2205
         Width           =   1155
      End
      Begin VB.TextBox DetOBS 
         Height          =   780
         Left            =   900
         MaxLength       =   250
         TabIndex        =   16
         Top             =   1380
         Width           =   9855
      End
      Begin VB.CommandButton BotaoHistorico 
         Caption         =   "Histórico"
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
         Left            =   8597
         TabIndex        =   24
         Top             =   2205
         Width           =   1155
      End
      Begin VB.CommandButton BotaoGravar 
         Caption         =   "Gravar"
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
         Left            =   9780
         TabIndex        =   25
         Top             =   2205
         Width           =   960
      End
      Begin VB.ComboBox DetAndamento 
         Height          =   315
         Left            =   5745
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   600
         Width           =   5010
      End
      Begin VB.CommandButton BotaoCR 
         Caption         =   "Titulos a Receber"
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
         Left            =   4724
         TabIndex        =   21
         Top             =   2205
         Width           =   1740
      End
      Begin VB.CommandButton BotaoCliente 
         Caption         =   "Cliente"
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
         Left            =   6490
         TabIndex        =   22
         Top             =   2205
         Width           =   900
      End
      Begin VB.CommandButton BotaoNF 
         Caption         =   "NFs"
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
         Left            =   3528
         TabIndex        =   20
         Top             =   2205
         Width           =   1170
      End
      Begin VB.CommandButton BotaoOP 
         Caption         =   "OPs"
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
         Left            =   2332
         TabIndex        =   19
         Top             =   2205
         Width           =   1170
      End
      Begin VB.CommandButton BotaoPV 
         Caption         =   "PV"
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
         Left            =   60
         TabIndex        =   17
         Top             =   2205
         Width           =   1170
      End
      Begin MSComCtl2.UpDown UpDownDetEntrega 
         Height          =   300
         Left            =   4125
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   615
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DetEntrega 
         Height          =   315
         Left            =   3000
         TabIndex        =   13
         Top             =   615
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label DetFilial 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9510
         TabIndex        =   61
         Top             =   210
         Width           =   1230
      End
      Begin VB.Label DetCliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5745
         TabIndex        =   60
         Top             =   210
         Width           =   3150
      End
      Begin VB.Label DetEmissao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3000
         TabIndex        =   59
         Top             =   225
         Width           =   1395
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "OBS:"
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
         Left            =   420
         TabIndex        =   58
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label DetFilialNF 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5745
         TabIndex        =   57
         Top             =   990
         Width           =   1440
      End
      Begin VB.Label DetNF 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3000
         TabIndex        =   56
         Top             =   1005
         Width           =   1395
      End
      Begin VB.Label DetOP 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   900
         TabIndex        =   55
         Top             =   1005
         Width           =   1440
      End
      Begin VB.Label DetValor 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   900
         TabIndex        =   54
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label DetNumero 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   900
         TabIndex        =   53
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial NF:"
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
         Index           =   9
         Left            =   4935
         TabIndex        =   52
         Top             =   1050
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "NFs:"
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
         Left            =   2505
         TabIndex        =   45
         Top             =   1065
         Width           =   405
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "OPs:"
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
         Left            =   465
         TabIndex        =   44
         Top             =   1050
         Width           =   420
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Andamento:"
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
         Index           =   8
         Left            =   4665
         TabIndex        =   43
         Top             =   645
         Width           =   1020
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Entrega:"
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
         Index           =   5
         Left            =   2235
         TabIndex        =   42
         Top             =   645
         Width           =   735
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Vlr Total:"
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
         Left            =   90
         TabIndex        =   41
         Top             =   660
         Width           =   795
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
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
         Index           =   7
         Left            =   9045
         TabIndex        =   40
         Top             =   255
         Width           =   465
      End
      Begin VB.Label label1 
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
         Height          =   195
         Index           =   6
         Left            =   5025
         TabIndex        =   39
         Top             =   240
         Width           =   660
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   0
         Left            =   2190
         TabIndex        =   38
         Top             =   255
         Width           =   750
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Pedido:"
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
         Left            =   225
         TabIndex        =   37
         Top             =   285
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pedidos de Venda que atendem aos Filtros"
      Height          =   2850
      Left            =   75
      TabIndex        =   30
      Top             =   1335
      Width           =   10830
      Begin VB.TextBox PVEmissao 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   6990
         TabIndex        =   65
         Top             =   2055
         Width           =   990
      End
      Begin VB.TextBox PVAndamento 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   8190
         TabIndex        =   51
         Top             =   2055
         Width           =   2475
      End
      Begin VB.TextBox PVEntrega 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   5940
         TabIndex        =   50
         Top             =   2055
         Width           =   990
      End
      Begin VB.TextBox PVValor 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   4845
         TabIndex        =   49
         Top             =   2055
         Width           =   1050
      End
      Begin VB.TextBox PVFilial 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   3615
         TabIndex        =   48
         Top             =   2055
         Width           =   1185
      End
      Begin VB.TextBox PVCliente 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   1560
         TabIndex        =   47
         Top             =   2055
         Width           =   2010
      End
      Begin VB.TextBox PVCodigo 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   630
         TabIndex        =   46
         Top             =   2055
         Width           =   885
      End
      Begin MSFlexGridLib.MSFlexGrid GridPV 
         Height          =   1680
         Left            =   60
         TabIndex        =   12
         Top             =   210
         Width           =   10650
         _ExtentX        =   18785
         _ExtentY        =   2963
         _Version        =   393216
         Rows            =   21
         Cols            =   8
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      Height          =   1290
      Left            =   75
      TabIndex        =   29
      Top             =   30
      Width           =   9585
      Begin VB.CheckBox SoAbertos 
         Caption         =   "Só trazer PVs abertos"
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
         Left            =   6105
         TabIndex        =   2
         Top             =   300
         Width           =   2280
      End
      Begin VB.Frame Frame4 
         Caption         =   "Período de Entrega"
         Height          =   615
         Left            =   4560
         TabIndex        =   62
         Top             =   570
         Width           =   3900
         Begin MSComCtl2.UpDown UpDownEntregaDe 
            Height          =   315
            Left            =   1350
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   225
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox EntregaDe 
            Height          =   300
            Left            =   405
            TabIndex        =   7
            Top             =   240
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownEntregaAte 
            Height          =   315
            Left            =   3405
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox EntregaAte 
            Height          =   300
            Left            =   2460
            TabIndex        =   9
            Top             =   255
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label label1 
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
            Height          =   240
            Index           =   11
            Left            =   90
            TabIndex        =   64
            Top             =   285
            Width           =   345
         End
         Begin VB.Label label1 
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
            Index           =   10
            Left            =   2085
            TabIndex        =   63
            Top             =   300
            Width           =   360
         End
      End
      Begin VB.Frame FrameData 
         Caption         =   "Período de Emissão"
         Height          =   615
         Left            =   465
         TabIndex        =   32
         Top             =   570
         Width           =   3990
         Begin MSComCtl2.UpDown UpDownEmissaoDe 
            Height          =   315
            Left            =   1350
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   225
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox EmissaoDe 
            Height          =   300
            Left            =   405
            TabIndex        =   3
            Top             =   240
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownEmissaoAte 
            Height          =   315
            Left            =   3480
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox EmissaoAte 
            Height          =   300
            Left            =   2535
            TabIndex        =   5
            Top             =   255
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label label1 
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
            Index           =   12
            Left            =   2160
            TabIndex        =   34
            Top             =   300
            Width           =   360
         End
         Begin VB.Label label1 
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
            Height          =   240
            Index           =   13
            Left            =   90
            TabIndex        =   33
            Top             =   285
            Width           =   345
         End
      End
      Begin VB.CommandButton BotaoTrazerPV 
         Caption         =   "Trazer PVs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   8640
         TabIndex        =   11
         Top             =   195
         Width           =   765
      End
      Begin MSMask.MaskEdBox Cliente 
         Height          =   300
         Left            =   2985
         TabIndex        =   1
         Top             =   210
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   300
         Left            =   885
         TabIndex        =   0
         Top             =   225
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin VB.Label LabelCliente 
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
         Height          =   195
         Left            =   2310
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   36
         Top             =   255
         Width           =   660
      End
      Begin VB.Label LabelCodigo 
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
         Height          =   195
         Left            =   135
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   35
         Top             =   270
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   9810
      ScaleHeight     =   450
      ScaleWidth      =   1050
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   120
      Width           =   1110
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   45
         Picture         =   "PVConsulta2.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   525
         Picture         =   "PVConsulta2.ctx":0532
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
End
Attribute VB_Name = "PVConsulta2Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlteradoSel As Integer
Dim iAlterado As Integer
Dim iLinhaAnt As Integer
Dim bDesabilitaCmdGridAux As Boolean
Dim bTrazendoDados As Boolean

Dim gcolPVs As New Collection
Dim gcolcolNFs As New Collection
Dim gcolcolOPs As New Collection
Dim gcolCli As New Collection
Dim gcolFilial As New Collection

Dim objGridPV As AdmGrid
Dim iGrid_PVCodigo_Col As Integer
Dim iGrid_PVCliente_Col As Integer
Dim iGrid_PVFilial_Col As Integer
Dim iGrid_PVValor_Col As Integer
Dim iGrid_PVEntrega_Col As Integer
Dim iGrid_PVEmissao_Col As Integer
Dim iGrid_PVAndamento_Col As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

Dim sClienteAnt As String

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Acompanhamento de Vendas"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "PVConsulta2"

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

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoCodigo = Nothing
    Set objEventoCliente = Nothing
    
    Set gcolPVs = Nothing
    Set gcolcolNFs = Nothing
    Set gcolcolOPs = Nothing
    Set gcolCli = Nothing
    Set gcolFilial = Nothing

    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205293)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    bDesabilitaCmdGridAux = False
    bTrazendoDados = False

    Set objEventoCodigo = New AdmEvento
    Set objEventoCliente = New AdmEvento
    
    lErro = CF("Carrega_Combo", DetAndamento, "PVAndamento", "Codigo", TIPO_INT, "Descricao", TIPO_STR)
    If lErro <> SUCESSO Then gError 205708

    lErro = Inicializa_GridPV(objGridPV)
    If lErro <> SUCESSO Then gError 205709

    iAlteradoSel = 0
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 205708, 205709

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205710)

    End Select

    iAlteradoSel = 0
    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
    
    bDesabilitaCmdGridAux = False
    bTrazendoDados = False
    
    iAlteradoSel = 0
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205711)

    End Select

    iAlteradoSel = 0
    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(ByVal objPV As ClassPedidoDeVenda) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objPV.lCodigo = StrParaLong(DetNumero.Caption)
    objPV.iFilialEmpresa = giFilialEmpresa
    
    objPV.iAndamento = Codigo_Extrai(DetAndamento.Text)
    objPV.sObs = DetOBS.Text
    objPV.dtDataEntrega = StrParaDate(DetEntrega.Text)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205712)

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
Dim objPV As New ClassPedidoDeVenda

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Se não tiver linha selecionada => Erro
    If GridPV.Row = 0 Then gError 205713

    'Preenche o objVeiculos
    lErro = Move_Tela_Memoria(objPV)
    If lErro <> SUCESSO Then gError 205714

    'Grava o/a Veiculos no Banco de Dados
    lErro = CF("PV_Grava_Andamento", objPV)
    If lErro <> SUCESSO Then gError 205715
    
    gcolPVs.Item(GridPV.Row).sObs = objPV.sObs
    gcolPVs.Item(GridPV.Row).iAndamento = objPV.iAndamento
    gcolPVs.Item(GridPV.Row).dtDataEntrega = objPV.dtDataEntrega
    GridPV.TextMatrix(GridPV.Row, iGrid_PVAndamento_Col) = DetAndamento.Text
    
    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 205713
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 205714, 205715

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205716)

    End Select

    Exit Function

End Function

Function Limpa_Tela_PV() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_PV

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    SoAbertos.Value = vbUnchecked
    DetAndamento.ListIndex = -1
    
    Call Grid_Limpa(objGridPV)
    
    DetNumero.Caption = ""
    DetEmissao.Caption = ""
    DetCliente.Caption = ""
    DetFilial.Caption = ""
    DetValor.Caption = ""
    DetOP.Caption = ""
    DetNF.Caption = ""
    DetFilialNF.Caption = ""
    
    Set gcolPVs = New Collection
    Set gcolcolNFs = New Collection
    Set gcolcolOPs = New Collection
    Set gcolCli = New Collection
    Set gcolFilial = New Collection
    
    sClienteAnt = ""

    iAlteradoSel = 0
    iAlterado = 0

    Limpa_Tela_PV = SUCESSO

    Exit Function

Erro_Limpa_Tela_PV:

    Limpa_Tela_PV = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205717)

    End Select

    Exit Function

End Function

Function Traz_PV_Tela(ByVal colPVs As Collection) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objPV As ClassPedidoDeVenda
Dim iIndice As Integer
Dim sAndamento As String
Dim objFilial As ClassFilialCliente
Dim objCli As ClassCliente

On Error GoTo Erro_Traz_PV_Tela

    Call Grid_Limpa(objGridPV)

    Set gcolPVs = colPVs
    Set gcolCli = New Collection
    Set gcolFilial = New Collection
    
    'Aumenta o número de linhas do grid se necessário
    If colPVs.Count >= objGridPV.objGrid.Rows Then
        Call Refaz_Grid(objGridPV, colPVs.Count)
    End If

    iLinha = 0
    For Each objPV In colPVs
    
        Set objCli = New ClassCliente
        Set objFilial = New ClassFilialCliente
    
        objCli.lCodigo = objPV.lCliente
        lErro = CF("Cliente_Le", objCli)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 205718
        
        objFilial.lCodCliente = objPV.lCliente
        objFilial.iCodFilial = objPV.iFilial
        lErro = CF("FilialCliente_Le", objFilial)
        If lErro <> SUCESSO And lErro <> 12567 Then gError 205719
        
        gcolCli.Add objCli
        gcolFilial.Add objFilial
    
        iLinha = iLinha + 1
        sAndamento = ""
        
        For iIndice = 0 To DetAndamento.ListCount - 1
            If Codigo_Extrai(DetAndamento.List(iIndice)) = objPV.iAndamento Then
                sAndamento = DetAndamento.List(iIndice)
                Exit For
            End If
        Next
    
        GridPV.TextMatrix(iLinha, iGrid_PVAndamento_Col) = sAndamento
        GridPV.TextMatrix(iLinha, iGrid_PVCliente_Col) = CStr(objCli.lCodigo) & SEPARADOR & objCli.sNomeReduzido
        GridPV.TextMatrix(iLinha, iGrid_PVCodigo_Col) = objPV.lCodigo
        If objPV.dtDataEmissao <> DATA_NULA Then GridPV.TextMatrix(iLinha, iGrid_PVEmissao_Col) = Format(objPV.dtDataEmissao, "dd/mm/yyyy")
        If objPV.dtDataEntrega <> DATA_NULA Then GridPV.TextMatrix(iLinha, iGrid_PVEntrega_Col) = Format(objPV.dtDataEntrega, "dd/mm/yyyy")
        GridPV.TextMatrix(iLinha, iGrid_PVFilial_Col) = CStr(objFilial.iCodFilial) & SEPARADOR & objFilial.sNome
        GridPV.TextMatrix(iLinha, iGrid_PVValor_Col) = Format(objPV.dValorTotal, "STANDARD")
    
    Next
    
    objGridPV.iLinhasExistentes = iLinha

    iAlteradoSel = 0

    Traz_PV_Tela = SUCESSO

    Exit Function

Erro_Traz_PV_Tela:

    Traz_PV_Tela = gErr

    Select Case gErr

        Case 205718, 205719

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205720)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 205721

    Call Rotina_Aviso(vbOKOnly, "AVISO_OPERACAO_SUCESSO")

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 205721

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205722)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205723)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 205724

    Call Limpa_Tela_PV

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 205724

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205725)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_Codigo_Validate

    'Verifica se Codigo está preenchida
    If Len(Trim(Codigo.Text)) <> 0 Then

        objPedidoVenda.lCodigo = StrParaLong(Codigo.Text)
        objPedidoVenda.iFilialEmpresa = giFilialEmpresa
   
        'Busca o pedido na tabela de Pedidos de Venda
        lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
        If lErro <> SUCESSO And lErro <> 26509 Then gError 205725
        If lErro <> SUCESSO Then
        
            'Verifica se o pedido está baixado
            lErro = CF("PedidoVendaBaixado_Le", objPedidoVenda)
            If lErro <> SUCESSO And lErro <> 46135 Then gError 205726
            If lErro = SUCESSO Then gError 205727
        
        End If
    
       Call BotaoTrazerPV_Click

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr
    
        Case 205725, 205726

        Case 205727
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO1", gErr, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205728)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlteradoSel)
    
End Sub

Private Sub Codigo_Change()
    iAlteradoSel = REGISTRO_ALTERADO
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPV As ClassPedidoDeVenda

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objPV = obj1

    Codigo.PromptInclude = False
    Codigo.Text = objPV.lCodigo
    Codigo.PromptInclude = True
    Call Codigo_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205729)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objPV As New ClassPedidoDeVenda
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then
        objPV.lCodigo = StrParaLong(Codigo.Text)
    End If

    Call Chama_Tela("PedidoVendaTodosLista", colSelecao, objPV, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205730)

    End Select

    Exit Sub

End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        End If
    
    End If

End Sub

Private Function Inicializa_GridPV(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Pedido")
    objGrid.colColuna.Add ("Cliente")
    objGrid.colColuna.Add ("Filial")
    objGrid.colColuna.Add ("Emissão")
    objGrid.colColuna.Add ("Entrega")
    objGrid.colColuna.Add ("Valor")
    objGrid.colColuna.Add ("Andamento")

    'Controles que participam do Grid
    objGrid.colCampo.Add (PVCodigo.Name)
    objGrid.colCampo.Add (PVCliente.Name)
    objGrid.colCampo.Add (PVFilial.Name)
    objGrid.colCampo.Add (PVEmissao.Name)
    objGrid.colCampo.Add (PVEntrega.Name)
    objGrid.colCampo.Add (PVValor.Name)
    objGrid.colCampo.Add (PVAndamento.Name)

    'Colunas do Grid
    iGrid_PVCodigo_Col = 1
    iGrid_PVCliente_Col = 2
    iGrid_PVFilial_Col = 3
    iGrid_PVEmissao_Col = 4
    iGrid_PVEntrega_Col = 5
    iGrid_PVValor_Col = 6
    iGrid_PVAndamento_Col = 7

    objGrid.objGrid = GridPV

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 500 + 1

    objGrid.iExecutaRotinaEnable = GRID_NAO_EXECUTAR_ROTINA_ENABLE
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    objGrid.iLinhasVisiveis = 9

    'Largura da primeira coluna
    GridPV.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridPV = SUCESSO

End Function

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim bCancel As Boolean

    Set objcliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objcliente.sNomeReduzido

    'Executa o Validate
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelCliente_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objcliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente)


End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer

On Error GoTo Erro_Cliente_Validate

    'Verifica se o Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 And sClienteAnt <> Cliente.Text Then

        'Busca o Cliente no BD
        lErro = TP_Cliente_Le2(Cliente, objcliente, iCodFilial)
        If lErro <> SUCESSO Then gError 205731

    End If
    
    sClienteAnt = Cliente.Text

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 205731
            Call Rotina_Erro(vbOKOnly, "Erro na validação do cliente.", gErr, Error, 205732)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205732)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_GotFocus()
    Call MaskEdBox_TrataGotFocus(Cliente, iAlteradoSel)
End Sub

Private Sub Cliente_Change()
    iAlteradoSel = REGISTRO_ALTERADO
End Sub

Private Sub EmissaoDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(EmissaoDe, iAlteradoSel)
End Sub

Private Sub EmissaoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoDe_Validate

    If Len(Trim(EmissaoDe.ClipText)) <> 0 Then

        lErro = Data_Critica(EmissaoDe.Text)
        If lErro <> SUCESSO Then gError 205733
    
    End If

    Exit Sub

Erro_EmissaoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 205733

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205734)

    End Select

    Exit Sub

End Sub

Private Sub EmissaoDe_Change()
    iAlteradoSel = REGISTRO_ALTERADO
End Sub

Private Sub UpDownEmissaoDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissaoDe_DownClick

    EmissaoDe.SetFocus

    If Len(EmissaoDe.ClipText) > 0 Then

        sData = EmissaoDe.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 205735

        EmissaoDe.Text = sData
        
        Call EmissaoDe_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownEmissaoDe_DownClick:

    Select Case gErr

        Case 205735

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205736)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    EmissaoDe.SetFocus

    If Len(Trim(EmissaoDe.ClipText)) > 0 Then

        sData = EmissaoDe.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 205737

        EmissaoDe.Text = sData
        
        Call EmissaoDe_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case gErr

        Case 205737

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205738)

    End Select

    Exit Sub

End Sub

Private Sub EmissaoAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(EmissaoAte, iAlteradoSel)
End Sub

Private Sub EmissaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoAte_Validate

    If Len(Trim(EmissaoAte.ClipText)) <> 0 Then

        lErro = Data_Critica(EmissaoAte.Text)
        If lErro <> SUCESSO Then gError 205739
    
    End If

    Exit Sub

Erro_EmissaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 205739

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205740)

    End Select

    Exit Sub

End Sub

Private Sub EmissaoAte_Change()
    iAlteradoSel = REGISTRO_ALTERADO
End Sub

Private Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    EmissaoAte.SetFocus

    If Len(EmissaoAte.ClipText) > 0 Then

        sData = EmissaoAte.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 205741

        EmissaoAte.Text = sData
        
        Call EmissaoAte_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownEmissaoAte_DownClick:

    Select Case gErr

        Case 205741

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205742)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissaoAte_UpClick

    EmissaoAte.SetFocus

    If Len(Trim(EmissaoAte.ClipText)) > 0 Then

        sData = EmissaoAte.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 205743

        EmissaoAte.Text = sData
        
        Call EmissaoAte_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownEmissaoAte_UpClick:

    Select Case gErr

        Case 205743

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205744)

    End Select

    Exit Sub

End Sub

Private Sub EntregaDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(EntregaDe, iAlteradoSel)
End Sub

Private Sub EntregaDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EntregaDe_Validate

    If Len(Trim(EntregaDe.ClipText)) <> 0 Then

        lErro = Data_Critica(EntregaDe.Text)
        If lErro <> SUCESSO Then gError 205745
    
    End If

    Exit Sub

Erro_EntregaDe_Validate:

    Cancel = True

    Select Case gErr

        Case 205745

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205746)

    End Select

    Exit Sub

End Sub

Private Sub EntregaDe_Change()
    iAlteradoSel = REGISTRO_ALTERADO
End Sub

Private Sub UpDownEntregaDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntregaDe_DownClick

    EntregaDe.SetFocus

    If Len(EntregaDe.ClipText) > 0 Then

        sData = EntregaDe.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 205747

        EntregaDe.Text = sData
        
        Call EntregaDe_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownEntregaDe_DownClick:

    Select Case gErr

        Case 205747

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205748)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntregaDe_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntregaDe_UpClick

    EntregaDe.SetFocus

    If Len(Trim(EntregaDe.ClipText)) > 0 Then

        sData = EntregaDe.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 205749

        EntregaDe.Text = sData
        
        Call EntregaDe_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownEntregaDe_UpClick:

    Select Case gErr

        Case 205749

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205750)

    End Select

    Exit Sub

End Sub

Private Sub EntregaAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(EntregaAte, iAlteradoSel)
End Sub

Private Sub EntregaAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EntregaAte_Validate

    If Len(Trim(EntregaAte.ClipText)) <> 0 Then

        lErro = Data_Critica(EntregaAte.Text)
        If lErro <> SUCESSO Then gError 205751
    
    End If

    Exit Sub

Erro_EntregaAte_Validate:

    Cancel = True

    Select Case gErr

        Case 205751

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205752)

    End Select

    Exit Sub

End Sub

Private Sub EntregaAte_Change()
    iAlteradoSel = REGISTRO_ALTERADO
End Sub

Private Sub UpDownEntregaAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntregaAte_DownClick

    EntregaAte.SetFocus

    If Len(EntregaAte.ClipText) > 0 Then

        sData = EntregaAte.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 205753

        EntregaAte.Text = sData
        
        Call EntregaAte_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownEntregaAte_DownClick:

    Select Case gErr

        Case 205753

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205754)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntregaAte_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntregaAte_UpClick

    EntregaAte.SetFocus

    If Len(Trim(EntregaAte.ClipText)) > 0 Then

        sData = EntregaAte.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 205755

        EntregaAte.Text = sData
        
        Call EntregaAte_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownEntregaAte_UpClick:

    Select Case gErr

        Case 205755

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205756)

    End Select

    Exit Sub

End Sub

Private Sub DetEntrega_GotFocus()
    Call MaskEdBox_TrataGotFocus(DetEntrega, iAlteradoSel)
End Sub

Private Sub DetEntrega_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DetEntrega_Validate

    If Len(Trim(DetEntrega.ClipText)) <> 0 Then

        lErro = Data_Critica(DetEntrega.Text)
        If lErro <> SUCESSO Then gError 205757
    
    End If

    Exit Sub

Erro_DetEntrega_Validate:

    Cancel = True

    Select Case gErr

        Case 205757

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205758)

    End Select

    Exit Sub

End Sub

Private Sub DetEntrega_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDetEntrega_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDetEntrega_DownClick

    DetEntrega.SetFocus

    If Len(DetEntrega.ClipText) > 0 Then

        sData = DetEntrega.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 205759

        DetEntrega.Text = sData
        
        Call DetEntrega_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownDetEntrega_DownClick:

    Select Case gErr

        Case 205759

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205760)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDetEntrega_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDetEntrega_UpClick

    DetEntrega.SetFocus

    If Len(Trim(DetEntrega.ClipText)) > 0 Then

        sData = DetEntrega.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 205761

        DetEntrega.Text = sData
        
        Call DetEntrega_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownDetEntrega_UpClick:

    Select Case gErr

        Case 205761

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205762)

    End Select

    Exit Sub

End Sub

Private Sub BotaoTrazerPV_Click()

Dim lErro As Long
Dim objPVSel As New ClassPVAndamentoSel

On Error GoTo Erro_BotaoTrazerPV_Click

    objPVSel.dtEmissaoAte = StrParaDate(EmissaoAte.Text)
    objPVSel.dtEmissaoDe = StrParaDate(EmissaoDe.Text)
    objPVSel.dtEntregaAte = StrParaDate(EntregaAte.Text)
    objPVSel.dtEntregaDe = StrParaDate(EntregaDe.Text)
    objPVSel.lCliente = LCodigo_Extrai(Cliente.Text)
    objPVSel.lPedido = StrParaLong(Codigo.Text)
    
    If SoAbertos.Value = vbChecked Then
        objPVSel.iSoAbertos = MARCADO
    Else
        objPVSel.iSoAbertos = DESMARCADO
    End If
    
    lErro = CF("PV_Le_Andamento", objPVSel)
    If lErro <> SUCESSO Then gError 205763
    
    lErro = Traz_PV_Tela(objPVSel.colPVs)
    If lErro <> SUCESSO Then gError 205764
    
    Set gcolcolNFs = objPVSel.colcolNFs
    Set gcolcolOPs = objPVSel.colcolOPs

    Exit Sub

Erro_BotaoTrazerPV_Click:

    Select Case gErr
    
        Case 205763, 205764

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205765)

    End Select

    Exit Sub
    
End Sub

Private Sub GridPV_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridPV, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPV, iAlteradoSel)
    End If

End Sub

Private Sub GridPV_GotFocus()
    Call Grid_Recebe_Foco(objGridPV)
End Sub

Private Sub GridPV_EnterCell()
    If Not bDesabilitaCmdGridAux Then
        Call Grid_Entrada_Celula(objGridPV, iAlteradoSel)
    End If
End Sub

Private Sub GridPV_LeaveCell()
    If Not bDesabilitaCmdGridAux Then
        Call Saida_Celula(objGridPV)
    End If
End Sub

Private Sub GridPV_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridPV, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPV, iAlteradoSel)
    End If

End Sub

Private Sub GridPV_RowColChange()

    If Not bDesabilitaCmdGridAux Then

        Call Grid_RowColChange(objGridPV)
        
        Call Mostra_Dados(GridPV.Row)
                
    End If
    
    iLinhaAnt = GridPV.Row
    
End Sub

Private Sub GridPV_Scroll()
    Call Grid_Scroll(objGridPV)
End Sub

Private Sub GridPV_KeyDown(KeyCode As Integer, Shift As Integer)
    
Dim lErro As Long
Dim iItemAtual As Integer
Dim iLinhasExistentesAnt As Integer
Dim vbMsgRes As VbMsgBoxResult
    
On Error GoTo Erro_GridPV_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnt = objGridPV.iLinhasExistentes
    iItemAtual = GridPV.Row
        
    Call Grid_Trata_Tecla1(KeyCode, objGridPV)

    Exit Sub

Erro_GridPV_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205766)

    End Select

    Exit Sub
    
End Sub

Private Sub GridPV_LostFocus()
    Call Grid_Libera_Foco(objGridPV)
End Sub

Private Function Mostra_Dados(ByVal iLinha As Integer) As Long

Dim lErro As Long, iIndice As Integer
Dim objPV As ClassPedidoDeVenda
Dim objCli As ClassCliente
Dim objFilial As ClassFilialCliente
Dim colNFs As Collection
Dim colOPs As Collection
Dim sNF As String
Dim sOP As String
Dim sFilial As String
Dim objNF As ClassNFiscal
Dim objOP As ClassOrdemDeProducao

On Error GoTo Erro_Mostra_Dados
    
    If iLinhaAnt <> 0 Then Call Teste_Salva(Me, iAlterado)

    DetEntrega.PromptInclude = False
    DetEntrega.Text = ""
    DetEntrega.PromptInclude = True
    DetOBS.Text = ""
    DetAndamento.ListIndex = -1
    DetNumero.Caption = ""
    DetEmissao.Caption = ""
    DetCliente.Caption = ""
    DetFilial.Caption = ""
    DetValor.Caption = ""
    DetOP.Caption = ""
    DetNF.Caption = ""
    DetFilialNF.Caption = ""

    If iLinha <> 0 And gcolPVs.Count >= iLinha And Not bTrazendoDados Then
                
        Set objPV = gcolPVs.Item(iLinha)
        Set objCli = gcolCli.Item(iLinha)
        Set objFilial = gcolFilial.Item(iLinha)
        Set colNFs = gcolcolNFs.Item(iLinha)
        Set colOPs = gcolcolOPs.Item(iLinha)
        
        DetEntrega.PromptInclude = False
        If objPV.dtDataEntrega <> DATA_NULA Then DetEntrega.Text = Format(objPV.dtDataEntrega, "dd/mm/yy")
        DetEntrega.PromptInclude = True
        DetOBS.Text = objPV.sObs
        Call Combo_Seleciona_ItemData(DetAndamento, objPV.iAndamento)
        DetNumero.Caption = CStr(objPV.lCodigo)
        If objPV.dtDataEmissao <> DATA_NULA Then DetEmissao.Caption = Format(objPV.dtDataEmissao, "dd/mm/yyyy")
        DetCliente.Caption = CStr(objCli.lCodigo) & SEPARADOR & objCli.sNomeReduzido
        DetFilial.Caption = CStr(objFilial.iCodFilial) & SEPARADOR & objFilial.sNome
        DetValor.Caption = Format(objPV.dValorTotal, "STANDARD")
        
        sOP = ""
        iIndice = 0
        For Each objOP In colOPs
            iIndice = iIndice + 1
            If iIndice = 1 Then
                sOP = objOP.sCodigo
            ElseIf iIndice = colOPs.Count Then
                sOP = sOP & " e " & objOP.sCodigo
            Else
                sOP = sOP & ", " & objOP.sCodigo
            End If
        Next
        
        sNF = ""
        sFilial = ""
        iIndice = 0
        For Each objNF In colNFs
            iIndice = iIndice + 1
            If iIndice = 1 Then
                sNF = CStr(objNF.lNumNotaFiscal)
                sFilial = CStr(objNF.iFilialEmpresa)
            ElseIf iIndice = colNFs.Count Then
                sNF = sNF & " e " & CStr(objNF.lNumNotaFiscal)
                sFilial = sFilial & " e " & CStr(objNF.iFilialEmpresa)
            Else
                sNF = sNF & ", " & CStr(objNF.lNumNotaFiscal)
                sFilial = sFilial & ", " & CStr(objNF.iFilialEmpresa)
            End If
        Next
        
        DetOP.Caption = sOP
        DetNF.Caption = sNF
        DetFilialNF.Caption = sFilial
        
    End If
    
    iAlterado = 0
    
    Mostra_Dados = SUCESSO
    
    Exit Function

Erro_Mostra_Dados:

    Mostra_Dados = gErr

    Select Case gErr
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205767)

    End Select

    Exit Function

End Function

Private Sub BotaoPV_Click()

Dim lErro As Long
Dim objPV As New ClassPedidoDeVenda

On Error GoTo Erro_BotaoPV_Click
    
    'Se não tiver linha selecionada => Erro
    If GridPV.Row = 0 Then gError 205768
       
    objPV.lCodigo = StrParaLong(DetNumero.Caption)
    objPV.iFilialEmpresa = giFilialEmpresa
 
    'Chama a tela de PV
    Call Chama_Tela("PedidoVenda", objPV)

    Exit Sub

Erro_BotaoPV_Click:

    Select Case gErr

        Case 205768
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205769)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoCliente_Click()

Dim lErro As Long
Dim objCli As New ClassCliente

On Error GoTo Erro_BotaoCliente_Click
    
    'Se não tiver linha selecionada => Erro
    If GridPV.Row = 0 Then gError 205770
       
    objCli.lCodigo = LCodigo_Extrai(DetCliente.Caption)
 
    'Chama a tela de Clientes
    Call Chama_Tela("Clientes", objCli)

    Exit Sub

Erro_BotaoCliente_Click:

    Select Case gErr

        Case 205770
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205771)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoItens_Click()

Dim lErro As Long
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoItens_Click
    
    'Se não tiver linha selecionada => Erro
    If GridPV.Row = 0 Then gError 205772
       
    colSelecao.Add StrParaLong(DetNumero.Caption)
    colSelecao.Add giFilialEmpresa
 
    Call Chama_Tela("ItensPedidoDeVendaTodosLista", colSelecao, Nothing, Nothing, "CodPedido = ? AND FilialEmpresa = ?")

    Exit Sub

Erro_BotaoItens_Click:

    Select Case gErr

        Case 205772
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205773)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoOP_Click()

Dim lErro As Long
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoOP_Click
    
    'Se não tiver linha selecionada => Erro
    If GridPV.Row = 0 Then gError 205774
       
    colSelecao.Add StrParaLong(DetNumero.Caption)
    colSelecao.Add giFilialEmpresa
 
    Call Chama_Tela("ItensOPTodos_ProdLista", colSelecao, Nothing, Nothing, "CodPedido = ? AND FilialEmpresa = ?")

    Exit Sub

Erro_BotaoOP_Click:

    Select Case gErr

        Case 205774
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205775)

    End Select
    
    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'GridNF
        If objGridInt.objGrid.Name = GridPV.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

            End Select
                    
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 205776

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 205776
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205777)

    End Select

    Exit Function

End Function

Private Sub BotaoNF_Click()

Dim lErro As Long
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoNF_Click
    
    'Se não tiver linha selecionada => Erro
    If GridPV.Row = 0 Then gError 205778
       
    colSelecao.Add StrParaLong(DetNumero.Caption)
    colSelecao.Add giFilialEmpresa
 
    Call Chama_Tela("NFiscalSaidaTodasLista", colSelecao, Nothing, Nothing, "NumPedidoVenda = ? AND FilialPedido = ?")

    Exit Sub

Erro_BotaoNF_Click:

    Select Case gErr

        Case 205778
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205779)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoCR_Click()

Dim lErro As Long
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoCR_Click
    
    'Se não tiver linha selecionada => Erro
    If GridPV.Row = 0 Then gError 205780
       
    colSelecao.Add StrParaLong(DetNumero.Caption)
    colSelecao.Add giFilialEmpresa
 
    Call Chama_Tela("TitRecTodosTFLista", colSelecao, Nothing, Nothing, "NumIntDoc IN (SELECT NumIntDocCPR FROM Nfiscal WHERE ClasseDocCPR = 2 AND NumPedidoVenda = ? AND FilialPedido = ? )")

    Exit Sub

Erro_BotaoCR_Click:

    Select Case gErr

        Case 205780
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205781)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoAnotacao_Click()

Dim lErro As Long
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoAnotacao_Click
    
    'Se não tiver linha selecionada => Erro
    If GridPV.Row = 0 Then gError 205782
    
    colSelecao.Add ANOTACAO_ORIGEM_PEDIDOVENDA
    colSelecao.Add CStr(giFilialEmpresa) & "," & DetNumero.Caption
 
    Call Chama_Tela("AnotacoesLista", colSelecao, Nothing, Nothing, "Origem = ? AND ID = ?")

    Exit Sub

Erro_BotaoAnotacao_Click:

    Select Case gErr

        Case 205782
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205783)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoHistorico_Click()

Dim lErro As Long
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoHistorico_Click
    
    'Se não tiver linha selecionada => Erro
    If GridPV.Row = 0 Then gError 205784
    
    colSelecao.Add StrParaLong(DetNumero.Caption)
    colSelecao.Add giFilialEmpresa
 
    Call Chama_Tela("PVHistAndLista", colSelecao, Nothing, Nothing, "Codigo = ? AND FilialEmpresa = ?")

    Exit Sub

Erro_BotaoHistorico_Click:

    Select Case gErr

        Case 205784
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205785)

    End Select
    
    Exit Sub

End Sub

Private Sub DetAndamento_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DetOBS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub
