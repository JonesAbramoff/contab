VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RotasOcx 
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   10995
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8745
      Top             =   690
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   8850
      ScaleHeight     =   450
      ScaleWidth      =   2025
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   120
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "Rotas.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "Rotas.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "Rotas.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "Rotas.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin VB.Frame FrameTrecho 
      Caption         =   "Trecho"
      Height          =   1440
      Left            =   60
      TabIndex        =   49
      Top             =   5415
      Width           =   10860
      Begin VB.ComboBox TrechoMeio 
         Height          =   315
         Left            =   555
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1005
         Width           =   2070
      End
      Begin MSMask.MaskEdBox TrechoDistancia 
         Height          =   315
         Left            =   5865
         TabIndex        =   20
         Top             =   1020
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TrechoTempo 
         Height          =   315
         Left            =   8790
         TabIndex        =   21
         Top             =   1020
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "m"
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
         Index           =   13
         Left            =   4365
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   65
         Top             =   1065
         Width           =   150
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   12
         Left            =   2565
         TabIndex        =   64
         Top             =   1065
         Width           =   630
      End
      Begin VB.Label DistAte 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3240
         TabIndex        =   63
         Top             =   1020
         Width           =   1110
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "min"
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
         Index           =   9
         Left            =   9615
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   59
         Top             =   1065
         Width           =   390
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "m"
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
         Index           =   8
         Left            =   7005
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   58
         Top             =   1065
         Width           =   150
      End
      Begin VB.Label TrechoEndAte 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   5880
         TabIndex        =   57
         Top             =   270
         Width           =   4905
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tempo:"
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
         Index           =   5
         Left            =   8130
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   56
         Top             =   1065
         Width           =   645
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   4
         Left            =   5475
         TabIndex        =   55
         Top             =   315
         Width           =   375
      End
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
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   54
         Top             =   285
         Width           =   300
      End
      Begin VB.Label TrechoEndDe 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   555
         TabIndex        =   53
         Top             =   255
         Width           =   4905
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Distância:"
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
         Index           =   6
         Left            =   4890
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   52
         Top             =   1065
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Meio:"
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
         Index           =   7
         Left            =   15
         TabIndex        =   51
         Top             =   1065
         Width           =   525
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transferência\Quebras"
      Height          =   1335
      Index           =   0
      Left            =   7785
      TabIndex        =   46
      Top             =   945
      Width           =   3135
      Begin VB.CommandButton BotaoTransf 
         Caption         =   "Transferir"
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
         Left            =   1650
         TabIndex        =   26
         ToolTipText     =   "Transfere os clientes selecionados para outra rota. A rota atual permanecerá com os clientes não selecionados"
         Top             =   990
         Width           =   1335
      End
      Begin VB.CommandButton BotaoQuebrar 
         Caption         =   "Quebrar"
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
         Left            =   150
         TabIndex        =   25
         ToolTipText     =   "Cria uma nova rota com os clientes selecionados. A rota atual permanecerá com os clientes não selecionados."
         Top             =   990
         Width           =   1335
      End
      Begin VB.CommandButton BotaoProxNumTransf 
         Height          =   285
         Left            =   2655
         Picture         =   "Rotas.ctx":0994
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Numeração Automática"
         Top             =   630
         Width           =   300
      End
      Begin VB.ComboBox ChaveTransf 
         Height          =   315
         Left            =   825
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   240
         Width           =   2145
      End
      Begin MSMask.MaskEdBox CodigoTransf 
         Height          =   315
         Left            =   825
         TabIndex        =   23
         Top             =   615
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   4
         PromptChar      =   " "
      End
      Begin VB.Label LabelCodigoTransf 
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
         Left            =   90
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   48
         Top             =   645
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Chave:"
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
         Index           =   0
         Left            =   30
         TabIndex        =   47
         Top             =   285
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação"
      Height          =   900
      Left            =   60
      TabIndex        =   42
      Top             =   30
      Width           =   8595
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   5805
         Picture         =   "Rotas.ctx":0A7E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Numeração Automática"
         Top             =   210
         Width           =   300
      End
      Begin VB.CheckBox Ativo 
         Caption         =   "Ativo"
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
         Left            =   6195
         TabIndex        =   3
         Top             =   195
         Value           =   1  'Checked
         Width           =   810
      End
      Begin VB.ComboBox Chave 
         Height          =   315
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   195
         Width           =   2145
      End
      Begin VB.TextBox Descricao 
         Height          =   315
         Left            =   1035
         MaxLength       =   250
         TabIndex        =   4
         Top             =   525
         Width           =   7455
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   3975
         TabIndex        =   1
         Top             =   195
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   4
         PromptChar      =   " "
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
         Height          =   315
         Left            =   7110
         TabIndex        =   5
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   2
         Left            =   75
         TabIndex        =   45
         Top             =   555
         Width           =   915
      End
      Begin VB.Label LabelChave 
         Alignment       =   1  'Right Justify
         Caption         =   "Chave:"
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
         Left            =   180
         TabIndex        =   44
         Top             =   240
         Width           =   825
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
         Left            =   3000
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   43
         Top             =   225
         Width           =   945
      End
   End
   Begin VB.Frame FrameVend 
      Caption         =   "Vendedores"
      Height          =   1335
      Left            =   60
      TabIndex        =   38
      Top             =   945
      Width           =   7485
      Begin VB.CommandButton BotaoVendedores 
         Caption         =   "Vendedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   45
         TabIndex        =   7
         Top             =   1020
         Width           =   1335
      End
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   225
         Left            =   150
         TabIndex        =   39
         Top             =   810
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox VendTelefone1 
         Height          =   225
         Left            =   2685
         TabIndex        =   40
         Top             =   810
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox VendTelefone2 
         Height          =   225
         Left            =   4005
         TabIndex        =   41
         Top             =   810
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridVend 
         Height          =   660
         Left            =   30
         TabIndex        =   6
         Top             =   180
         Width           =   7410
         _ExtentX        =   13070
         _ExtentY        =   1164
         _Version        =   393216
         Rows            =   21
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
   End
   Begin VB.Frame FrameParadas 
      Caption         =   "Paradas"
      Height          =   3120
      Left            =   60
      TabIndex        =   32
      Top             =   2280
      Width           =   10860
      Begin VB.Frame Frame2 
         Caption         =   "Ordenação"
         Height          =   2940
         Index           =   1
         Left            =   9660
         TabIndex        =   66
         Top             =   120
         Width           =   1155
         Begin VB.Frame Frame3 
            Caption         =   "Troca"
            Height          =   1020
            Left            =   60
            TabIndex        =   67
            Top             =   1890
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
               TabIndex        =   18
               Top             =   675
               Width           =   960
            End
            Begin MSMask.MaskEdBox LinhaDesejada 
               Height          =   315
               Left            =   540
               TabIndex        =   17
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
               TabIndex        =   70
               Top             =   165
               Width           =   450
            End
            Begin VB.Label LinhaAtual 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   30
               TabIndex        =   69
               Top             =   345
               Width           =   480
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
               TabIndex        =   68
               Top             =   165
               Width           =   300
            End
         End
         Begin VB.CommandButton BotaoTopo 
            Height          =   315
            Left            =   315
            Picture         =   "Rotas.ctx":0B68
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   570
            Width           =   540
         End
         Begin VB.CommandButton BotaoDesce 
            Height          =   315
            Left            =   315
            Picture         =   "Rotas.ctx":0E7A
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   1230
            Width           =   540
         End
         Begin VB.CommandButton BotaoSobe 
            Height          =   315
            Left            =   315
            Picture         =   "Rotas.ctx":103C
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   900
            Width           =   540
         End
         Begin VB.CommandButton BotaoFundo 
            Height          =   315
            Left            =   315
            Picture         =   "Rotas.ctx":11FE
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   1575
            Width           =   540
         End
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
            TabIndex        =   12
            Top             =   240
            Width           =   960
         End
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Height          =   330
         Left            =   1215
         Picture         =   "Rotas.ctx":1510
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2715
         Width           =   1095
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Height          =   330
         Left            =   2370
         Picture         =   "Rotas.ctx":252A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2715
         Width           =   1095
      End
      Begin MSMask.MaskEdBox ParadasFilial 
         Height          =   225
         Left            =   0
         TabIndex        =   60
         Top             =   1620
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   " "
      End
      Begin VB.CommandButton BotaoClientes 
         Caption         =   "Clientes"
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
         TabIndex        =   9
         Top             =   2715
         Width           =   1095
      End
      Begin VB.CheckBox ParadasSel 
         DragMode        =   1  'Automatic
         Height          =   225
         Left            =   525
         TabIndex        =   50
         Top             =   1230
         Width           =   375
      End
      Begin MSMask.MaskEdBox ParadasCliente 
         Height          =   225
         Left            =   1005
         TabIndex        =   33
         Top             =   1215
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ParadasEndereco 
         Height          =   225
         Left            =   4485
         TabIndex        =   34
         Top             =   1005
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ParadasBairro 
         Height          =   225
         Left            =   2955
         TabIndex        =   35
         Top             =   1215
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ParadasDistancia 
         Height          =   225
         Left            =   6705
         TabIndex        =   36
         Top             =   1215
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ParadasOBS 
         Height          =   225
         Left            =   5115
         TabIndex        =   37
         Top             =   1500
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridParadas 
         Height          =   900
         Left            =   15
         TabIndex        =   8
         Top             =   195
         Width           =   9570
         _ExtentX        =   16880
         _ExtentY        =   1588
         _Version        =   393216
         Rows            =   21
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label DistTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   7965
         TabIndex        =   62
         Top             =   2775
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Distância Total:"
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
         Left            =   6345
         TabIndex        =   61
         Top             =   2805
         Width           =   1545
      End
   End
End
Attribute VB_Name = "RotasOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim giContadorTempo As Integer
Dim giTentativa As Integer
Dim glNumIntRel As Long
Dim gsDiretorio As String

Dim iAlterado As Integer
Dim gobjRota As New ClassRotas
Dim gobjMeios As ClassCamposGenericos
Dim iLinhaAnt As Integer
Dim bDesabilitaCmdGridAux As Boolean
Dim bTrazendoDados As Boolean

Dim objGridVend As AdmGrid
Dim iGrid_Vendedor_Col As Integer
Dim iGrid_VendTelefone1_Col As Integer
Dim iGrid_VendTelefone2_Col As Integer

Dim objGridParadas As AdmGrid
Dim iGrid_ParadasSel_Col As Integer
Dim iGrid_ParadasCliente_Col As Integer
Dim iGrid_ParadasFilial_Col As Integer
Dim iGrid_ParadasEndereco_Col As Integer
Dim iGrid_ParadasBairro_Col As Integer
Dim iGrid_ParadasDistancia_Col As Integer
Dim iGrid_ParadasOBS_Col As Integer

Private WithEvents objEventoRota As AdmEvento
Attribute objEventoRota.VB_VarHelpID = -1
Private WithEvents objEventoRotaTransf As AdmEvento
Attribute objEventoRotaTransf.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Rotas"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Rotas"

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
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objGridVend = Nothing
    Set objGridParadas = Nothing
    
    Set gobjRota = Nothing
    Set gobjMeios = Nothing

    Set objEventoRota = Nothing
    Set objEventoCliente = Nothing
    Set objEventoVendedor = Nothing
    Set objEventoRotaTransf = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205148)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoRota = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    Set objEventoRotaTransf = New AdmEvento
    
    bDesabilitaCmdGridAux = False
    bTrazendoDados = False
    glNumIntRel = 0
    
    Set gobjMeios = New ClassCamposGenericos

    gobjMeios.lCodigo = CAMPOSGENERICOS_MEIOS_TRANSP
    
    lErro = CF("CamposGenericosValores_Le_CodCampo", gobjMeios)
    If lErro <> SUCESSO Then gError 205151

    lErro = Inicializa_GridVend(objGridVend)
    If lErro <> SUCESSO Then gError 205149

    lErro = Inicializa_GridParadas(objGridParadas)
    If lErro <> SUCESSO Then gError 205150
    
    'Carrega a combo Tipo
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_CHAVE_ROTA, Chave)
    If lErro <> SUCESSO Then gError 205151
    
    'Carrega a combo Tipo
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_CHAVE_ROTA, ChaveTransf)
    If lErro <> SUCESSO Then gError 205152
    
    'Carrega a combo Tipo
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_MEIOS_TRANSP, TrechoMeio)
    If lErro <> SUCESSO Then gError 205153

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 205149 To 205153

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205154)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objRota As ClassRotas) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    bDesabilitaCmdGridAux = False

    If Not (objRota Is Nothing) Then

        lErro = Traz_Rotas_Tela(objRota)
        If lErro <> SUCESSO Then gError 205155

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 205155

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205156)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(ByVal objRota As ClassRotas, Optional ByVal objRotaTrans As ClassRotas, Optional ByVal bQuebraTransf As Boolean = False) As Long

Dim lErro As Long
Dim objRotaVend As ClassRotasVend
Dim objRotaPontos As ClassRotasPontos
Dim iIndice As Integer
Dim iAnt As Integer

On Error GoTo Erro_Move_Tela_Memoria

    Call Recolhe_Dados(iLinhaAnt)

    objRota.sCodigo = Codigo.Text
    If Len(Trim(Chave)) > 0 Then objRota.lChave = Chave.ItemData(Chave.ListIndex)
    objRota.iFilialEmpresa = giFilialEmpresa
    objRota.sDescricao = Descricao.Text
    
    objRota.iAtivo = DESMARCADO
    If Ativo.Value = vbChecked Then
        objRota.iAtivo = MARCADO
    End If
    
    If bQuebraTransf Then
    
        objRotaTrans.sCodigo = CodigoTransf.Text
        objRotaTrans.lChave = ChaveTransf.ItemData(ChaveTransf.ListIndex)
        objRotaTrans.iFilialEmpresa = objRota.iFilialEmpresa
        objRotaTrans.sDescricao = objRota.sDescricao
        objRotaTrans.iAtivo = MARCADO
    
    End If
    
    For iIndice = 1 To objGridVend.iLinhasExistentes
        Set objRotaVend = New ClassRotasVend
        objRotaVend.iVendedor = Codigo_Extrai(GridVend.TextMatrix(iIndice, iGrid_Vendedor_Col))
        objRotaVend.iSeq = iIndice
        objRota.colVend.Add objRotaVend
        If bQuebraTransf Then
            Set objRotaVend = New ClassRotasVend
            objRotaVend.iVendedor = Codigo_Extrai(GridVend.TextMatrix(iIndice, iGrid_Vendedor_Col))
            objRotaVend.iSeq = iIndice
            objRotaTrans.colVend.Add objRotaVend
        End If
    Next
    
    iAnt = 0
    For iIndice = 1 To objGridParadas.iLinhasExistentes
        Set objRotaPontos = New ClassRotasPontos
        objRotaPontos.lCliente = LCodigo_Extrai(GridParadas.TextMatrix(iIndice, iGrid_ParadasCliente_Col))
        objRotaPontos.iFilialCliente = Codigo_Extrai(GridParadas.TextMatrix(iIndice, iGrid_ParadasFilial_Col))
        objRotaPontos.sObservacao = GridParadas.TextMatrix(iIndice, iGrid_ParadasOBS_Col)
        objRotaPontos.dDistancia = gobjRota.colPontos.Item(iIndice).dDistancia
        objRotaPontos.dTempo = gobjRota.colPontos.Item(iIndice).dTempo
        objRotaPontos.lMeio = gobjRota.colPontos.Item(iIndice).lMeio
        objRotaPontos.iSelecionado = StrParaInt(GridParadas.TextMatrix(iIndice, iGrid_ParadasSel_Col))
        If Not bQuebraTransf Or objRotaPontos.iSelecionado = DESMARCADO Then
            objRotaPontos.iSeq = objRota.colPontos.Count + 1
            'Limpa os dados do trecho porque o ponto anterios não vai para essa coleção
            If iAnt <> 1 Then
                objRotaPontos.dDistancia = 0
                objRotaPontos.dTempo = 0
                objRotaPontos.lMeio = 0
            End If
            objRota.colPontos.Add objRotaPontos
            iAnt = 1
        Else
            objRotaPontos.iSeq = objRotaTrans.colPontos.Count + 1
            'Limpa os dados do trecho porque o ponto anterios não vai para essa coleção
            If iAnt <> 2 Then
                objRotaPontos.dDistancia = 0
                objRotaPontos.dTempo = 0
                objRotaPontos.lMeio = 0
            End If
            objRotaTrans.colPontos.Add objRotaPontos
            iAnt = 2
        End If
    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205157)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objRota As New ClassRotas

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Rotas"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objRota)
    If lErro <> SUCESSO Then gError 205159

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objRota.sCodigo, STRING_ROTA_CODIGO, "Codigo"
    colCampoValor.Add "Chave", objRota.lChave, 0, "Chave"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 205159

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205160)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objRota As New ClassRotas

On Error GoTo Erro_Tela_Preenche

    objRota.sCodigo = colCampoValor.Item("Codigo").vValor
    objRota.lChave = colCampoValor.Item("Chave").vValor

    If objRota.sCodigo <> "" And objRota.lChave <> 0 Then

        lErro = Traz_Rotas_Tela(objRota)
        If lErro <> SUCESSO Then gError 205161

    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 205161

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205162)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objRota As New ClassRotas

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Text)) = 0 Then gError 205163
    If Len(Trim(Chave.Text)) = 0 Then gError 205164

    'Preenche o objRota
    lErro = Move_Tela_Memoria(objRota)
    If lErro <> SUCESSO Then gError 205165
    
    If objRota.colPontos.Count = 0 Then gError 205166

    lErro = Trata_Alteracao(objRota, objRota.sCodigo, objRota.lChave)
    If lErro <> SUCESSO Then gError 205167

    'Grava o/a Rotas no Banco de Dados
    lErro = CF("Rotas_Grava", objRota)
    If lErro <> SUCESSO Then gError 205168

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 205163
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_ROTA_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 205164
            Call Rotina_Erro(vbOKOnly, "ERRO_CHAVE_ROTA_NAO_PREENCHIDO", gErr)
            Chave.SetFocus
            
        Case 205165, 205167, 205168
        
        Case 205166
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTAS_SEM_PARADAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205169)

    End Select

    Exit Function

End Function

Function Limpa_Tela_Rotas() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Rotas

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    glNumIntRel = 0
    
    TrechoEndDe.Caption = ""
    TrechoEndAte.Caption = ""
    TrechoMeio.ListIndex = -1
    Chave.ListIndex = -1
    ChaveTransf.ListIndex = -1
    Ativo.Value = vbChecked
    
    TrechoDistancia.Text = ""
    TrechoTempo.Text = ""
    TrechoMeio.ListIndex = -1
    TrechoEndAte.Caption = ""
    TrechoEndDe.Caption = ""
    FrameTrecho.Enabled = False
    FrameTrecho.Caption = "Trecho"
    DistAte.Caption = ""
       
    Set gobjRota = New ClassRotas

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    Call Grid_Limpa(objGridVend)
    Call Grid_Limpa(objGridParadas)

    iAlterado = 0

    Limpa_Tela_Rotas = SUCESSO

    Exit Function

Erro_Limpa_Tela_Rotas:

    Limpa_Tela_Rotas = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205170)

    End Select

    Exit Function

End Function

Function Traz_Rotas_Tela(objRota As ClassRotas) As Long

Dim lErro As Long
Dim objRotaVend As ClassRotasVend
Dim objRotaPontos As ClassRotasPontos
Dim objRotaPontosAnt As ClassRotasPontos
Dim iLinha As Integer
Dim objcliente As ClassCliente
Dim objVendedor As ClassVendedor
Dim objEndereco As ClassEndereco
Dim objFilial As ClassFilialCliente

On Error GoTo Erro_Traz_Rotas_Tela

    bTrazendoDados = True

    Call Limpa_Tela_Rotas
    
    Set gobjRota = objRota

    If objRota.sCodigo <> "" Then
        Codigo.Text = objRota.sCodigo
    End If

    If objRota.lChave <> 0 Then
        Call Combo_Seleciona_ItemData(Chave, objRota.lChave)
    End If

    'Lê o Rotas que está sendo Passado
    lErro = CF("Rotas_Le", objRota)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 205171

    If lErro = SUCESSO Then

        Descricao.Text = objRota.sDescricao

        If objRota.iAtivo = MARCADO Then
            Ativo.Value = vbChecked
        Else
            Ativo.Value = vbUnchecked
        End If
        
        iLinha = 0
        For Each objRotaVend In objRota.colVend
        
            Set objEndereco = New ClassEndereco
            Set objVendedor = New ClassVendedor
            
            iLinha = iLinha + 1
            Vendedor.Text = CStr(objRotaVend.iVendedor)
            Call TP_Vendedor_Le2(Vendedor, objVendedor)
            GridVend.TextMatrix(iLinha, iGrid_Vendedor_Col) = Vendedor.Text
        
            objEndereco.lCodigo = objVendedor.lEndereco
            
            'Le endereço no BD
            lErro = CF("Endereco_le", objEndereco)
            If lErro <> SUCESSO Then gError 205172
        
            GridVend.TextMatrix(iLinha, iGrid_VendTelefone1_Col) = objEndereco.sTelefone1
            GridVend.TextMatrix(iLinha, iGrid_VendTelefone2_Col) = objEndereco.sTelefone2
        
        Next
        objGridVend.iLinhasExistentes = iLinha
        
        lErro = Traz_Pontos_Tela(objRota)
        If lErro <> SUCESSO Then gError 205173

    End If
    
    Call Soma_Coluna_Grid(objGridParadas, iGrid_ParadasDistancia_Col, DistTotal, False)

    bTrazendoDados = False
    
    Call Mostra_Dados(GridParadas.Row)

    iAlterado = 0

    Traz_Rotas_Tela = SUCESSO

    Exit Function

Erro_Traz_Rotas_Tela:

    bTrazendoDados = False

    Traz_Rotas_Tela = gErr

    Select Case gErr

        Case 205171 To 205173

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205174)

    End Select

    Exit Function

End Function

Function Traz_Pontos_Tela(objRota As ClassRotas) As Long

Dim lErro As Long
Dim objRotaPontos As ClassRotasPontos
Dim objRotaPontosAnt As ClassRotasPontos
Dim iLinha As Integer
Dim objcliente As ClassCliente
Dim objEndereco As ClassEndereco
Dim objFilial As ClassFilialCliente
Dim objCamposGenericosValores As ClassCamposGenericosValores
Dim dDistAte As Double
Dim lEndereco2 As Long

On Error GoTo Erro_Traz_Pontos_Tela
        
    bTrazendoDados = True
        
    iLinha = 0
    For Each objRotaPontos In objRota.colPontos
    
        Set objEndereco = New ClassEndereco
        Set objcliente = New ClassCliente
        Set objFilial = New ClassFilialCliente
        
        iLinha = iLinha + 1
        ParadasCliente.Text = objRotaPontos.lCliente
        Call TP_Cliente_Le2(ParadasCliente, objcliente)
        GridParadas.TextMatrix(iLinha, iGrid_ParadasCliente_Col) = ParadasCliente.Text
        
        objEndereco.lCodigo = objcliente.lEnderecoEntrega
        lEndereco2 = objcliente.lEndereco
        
        If objRotaPontos.iFilialCliente <> 0 Then
        
            objFilial.lCodCliente = objRotaPontos.lCliente
            objFilial.iCodFilial = objRotaPontos.iFilialCliente

            lErro = CF("FilialCliente_Le", objFilial)
            If lErro <> SUCESSO And lErro <> 12567 Then gError 205173
            
            objEndereco.lCodigo = objFilial.lEnderecoEntrega
            lEndereco2 = objFilial.lEndereco
            
            GridParadas.TextMatrix(iLinha, iGrid_ParadasFilial_Col) = CStr(objFilial.iCodFilial) & SEPARADOR & objFilial.sNome
        
        End If
        
        'Le endereço no BD
        lErro = CF("Endereco_le", objEndereco)
        If lErro <> SUCESSO Then gError 205173
        
        If Len(Trim(objEndereco.sEndereco)) = 0 Then
        
            objEndereco.lCodigo = lEndereco2
        
            'Le endereço no BD
            lErro = CF("Endereco_le", objEndereco)
            If lErro <> SUCESSO Then gError 205173
        
        End If

        GridParadas.TextMatrix(iLinha, iGrid_ParadasEndereco_Col) = objEndereco.sEndereco
        GridParadas.TextMatrix(iLinha, iGrid_ParadasBairro_Col) = objEndereco.sBairro
        GridParadas.TextMatrix(iLinha, iGrid_ParadasOBS_Col) = objRotaPontos.sObservacao
        If objRotaPontos.dDistancia <> 0 Then GridParadas.TextMatrix(iLinha, iGrid_ParadasDistancia_Col) = Formata_Estoque(objRotaPontos.dDistancia)
        GridParadas.TextMatrix(iLinha, iGrid_ParadasSel_Col) = objRotaPontos.iSelecionado
                
        Set objRotaPontos.objcliente = objcliente
        Set objRotaPontos.objEndereco = objEndereco
        If iLinha <> 1 Then Set objRotaPontos.objPontoAnt = objRotaPontosAnt
        
        Set objRotaPontosAnt = objRotaPontos
    
    Next
    objGridParadas.iLinhasExistentes = iLinha
    
    Call Grid_Refresh_Checkbox(objGridParadas)
    
    iLinha = 0
    For Each objRotaPontos In objRota.colPontos
        iLinha = iLinha + 1
        dDistAte = 0
        If objRotaPontos.lMeio <> 0 Then
            For Each objCamposGenericosValores In gobjMeios.colCamposGenericosValores
                If objRotaPontos.lMeio = objCamposGenericosValores.lCodValor Then
                    If IsNumeric(objCamposGenericosValores.sComplemento1) Then
                        dDistAte = StrParaDbl(objCamposGenericosValores.sComplemento1)
                    End If
                    Exit For
                End If
            Next
        End If
                
        bDesabilitaCmdGridAux = True
        GridParadas.Row = iLinha
        GridParadas.Col = iGrid_ParadasDistancia_Col
        If objRotaPontos.dDistancia > dDistAte And dDistAte > 0 Then
            GridParadas.CellForeColor = vbRed
        Else
            GridParadas.CellForeColor = vbBlack
        End If
        bDesabilitaCmdGridAux = False
    Next
    
    bTrazendoDados = False
    
    iAlterado = 0

    Traz_Pontos_Tela = SUCESSO

    Exit Function

Erro_Traz_Pontos_Tela:

    bTrazendoDados = False

    bDesabilitaCmdGridAux = False

    Traz_Pontos_Tela = gErr

    Select Case gErr

        Case 205171 To 205173

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205174)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 205175

    'Limpa Tela
    Call Limpa_Tela_Rotas

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 205175

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205176)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205177)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 205178

    Call Limpa_Tela_Rotas

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 205178

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205179)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objRota As New ClassRotas
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Text)) = 0 Then gError 205180
    If Len(Trim(Chave.Text)) = 0 Then gError 205181

    objRota.sCodigo = Codigo.Text
    objRota.lChave = Chave.ItemData(Chave.ListIndex)
    objRota.iFilialEmpresa = giFilialEmpresa

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_ROTAS", objRota.sCodigo, objRota.lChave)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("Rotas_Exclui", objRota)
        If lErro <> SUCESSO Then gError 205182

        'Limpa Tela
        Call Limpa_Tela_Rotas

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 205180
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_ROTA_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 205181
            Call Rotina_Erro(vbOKOnly, "ERRO_CHAVE_ROTA_NAO_PREENCHIDO", gErr)
            Chave.SetFocus

        Case 205182

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205183)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate


    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205184)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Chave_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Chave_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Descricao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Descricao_Validate

    'Verifica se Descricao está preenchida
    If Len(Trim(Descricao.Text)) <> 0 Then

    End If

    Exit Sub

Erro_Descricao_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205185)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Ativo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoRota_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRota As ClassRotas

On Error GoTo Erro_objEventoRota_evSelecao

    Set objRota = obj1

    'Mostra os dados do Rotas na tela
    lErro = Traz_Rotas_Tela(objRota)
    If lErro <> SUCESSO Then gError 205186

    Me.Show

    Exit Sub

Erro_objEventoRota_evSelecao:

    Select Case gErr

        Case 205186

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205187)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objRota As New ClassRotas
Dim colSelecao As New Collection
Dim sFiltro As String

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objRota.sCodigo = Codigo.Text

    End If
    
    If Len(Trim(Chave.Text)) > 0 Then
        sFiltro = "ChaveCod = ?"
        colSelecao.Add Chave.ItemData(Chave.ListIndex)
    End If

    Call Chama_Tela("RotasLista", colSelecao, objRota, objEventoRota, sFiltro)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205188)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_GridVend(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Vendedor")
    objGrid.colColuna.Add ("Telefone 1")
    objGrid.colColuna.Add ("Telefone 2")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Vendedor.Name)
    objGrid.colCampo.Add (VendTelefone1.Name)
    objGrid.colCampo.Add (VendTelefone2.Name)

    'Colunas do Grid
    iGrid_Vendedor_Col = 1
    iGrid_VendTelefone1_Col = 2
    iGrid_VendTelefone2_Col = 3

    objGrid.objGrid = GridVend

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 10 + 1

    objGrid.iExecutaRotinaEnable = GRID_NAO_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 2

    'Largura da primeira coluna
    GridVend.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridVend = SUCESSO

End Function

Private Function Inicializa_GridParadas(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Cliente")
    objGrid.colColuna.Add ("Filial")
    objGrid.colColuna.Add ("S")
    objGrid.colColuna.Add ("Bairro")
    objGrid.colColuna.Add ("Endereço")
    objGrid.colColuna.Add ("Distância")
    objGrid.colColuna.Add ("Observação")

    'Controles que participam do Grid
    objGrid.colCampo.Add (ParadasCliente.Name)
    objGrid.colCampo.Add (ParadasFilial.Name)
    objGrid.colCampo.Add (ParadasSel.Name)
    objGrid.colCampo.Add (ParadasBairro.Name)
    objGrid.colCampo.Add (ParadasEndereco.Name)
    objGrid.colCampo.Add (ParadasDistancia.Name)
    objGrid.colCampo.Add (ParadasOBS.Name)

    'Colunas do Grid
    iGrid_ParadasCliente_Col = 1
    iGrid_ParadasFilial_Col = 2
    iGrid_ParadasSel_Col = 3
    iGrid_ParadasBairro_Col = 4
    iGrid_ParadasEndereco_Col = 5
    iGrid_ParadasDistancia_Col = 6
    iGrid_ParadasOBS_Col = 7

    objGrid.objGrid = GridParadas

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 200 + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridParadas.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridParadas = SUCESSO

End Function

Private Sub GridVend_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridVend, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridVend, iAlterado)
    End If

End Sub

Private Sub GridVend_GotFocus()
    Call Grid_Recebe_Foco(objGridVend)
End Sub

Private Sub GridVend_EnterCell()
    Call Grid_Entrada_Celula(objGridVend, iAlterado)
End Sub

Private Sub GridVend_LeaveCell()
    Call Saida_Celula(objGridVend)
End Sub

Private Sub GridVend_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridVend, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridVend, iAlterado)
    End If

End Sub

Private Sub GridVend_RowColChange()
    Call Grid_RowColChange(objGridVend)
End Sub

Private Sub GridVend_Scroll()
    Call Grid_Scroll(objGridVend)
End Sub

Private Sub GridVend_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridVend)
End Sub

Private Sub GridVend_LostFocus()
    Call Grid_Libera_Foco(objGridVend)
End Sub

Private Sub Vendedor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Vendedor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridVend)
End Sub

Private Sub Vendedor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridVend)
End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridVend.objControle = Vendedor
    lErro = Grid_Campo_Libera_Foco(objGridVend)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub VendTelefone1_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub VendTelefone1_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridVend)
End Sub

Private Sub VendTelefone1_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridVend)
End Sub

Private Sub VendTelefone1_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridVend.objControle = VendTelefone1
    lErro = Grid_Campo_Libera_Foco(objGridVend)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub VendTelefone2_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub VendTelefone2_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridVend)
End Sub

Private Sub VendTelefone2_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridVend)
End Sub

Private Sub VendTelefone2_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridVend.objControle = VendTelefone2
    lErro = Grid_Campo_Libera_Foco(objGridVend)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_Vendedor(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objEndereco As ClassEndereco
Dim objVendedor As ClassVendedor
            
On Error GoTo Erro_Saida_Celula_Vendedor

    Set objGridInt.objControle = Vendedor
    
    If Len(Trim(Vendedor.Text)) > 0 Then

        Set objEndereco = New ClassEndereco
        Set objVendedor = New ClassVendedor
        
        Vendedor.Text = LCodigo_Extrai(Vendedor.Text)
        
        'Verifica se Vendedor existe
        lErro = TP_Vendedor_Le2(Vendedor, objVendedor)
        If lErro <> SUCESSO And lErro <> 25018 And lErro <> 25020 Then gError 205189

        If lErro = 25018 Then gError 205190
        If lErro = 25020 Then gError 205191
    
        Call TP_Vendedor_Le2(Vendedor, objVendedor)
        GridVend.TextMatrix(GridVend.Row, iGrid_Vendedor_Col) = Vendedor.Text
    
        objEndereco.lCodigo = objVendedor.lEndereco
        
        'Le endereço no BD
        lErro = CF("Endereco_le", objEndereco)
        If lErro <> SUCESSO Then gError 205192
    
        GridVend.TextMatrix(GridVend.Row, iGrid_VendTelefone1_Col) = objEndereco.sTelefone1
        GridVend.TextMatrix(GridVend.Row, iGrid_VendTelefone2_Col) = objEndereco.sTelefone2

        If (GridVend.Row - GridVend.FixedRows) = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 205193

    Saida_Celula_Vendedor = SUCESSO

    Exit Function

Erro_Saida_Celula_Vendedor:

    Saida_Celula_Vendedor = gErr

    Select Case gErr

        Case 205189, 205192, 205193
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 205190 'Não encontrou nome reduzido de vendedor no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR")

            If vbMsgRes = vbYes Then

                'Preenche objVendedor com nome reduzido
                objVendedor.sNomeReduzido = Vendedor.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a tela de Vendedores
                Call Chama_Tela("Vendedores", objVendedor)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 205191 'Não encontrou codigo do vendedor no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR")

            If vbMsgRes = vbYes Then

                'Prenche objVendedor com codigo
                objVendedor.iCodigo = Codigo_Extrai(Vendedor.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a tela de Vendedores
                Call Chama_Tela("Vendedores", objVendedor)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205194)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Sub GridParadas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridParadas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParadas, iAlterado)
    End If

End Sub

Private Sub GridParadas_GotFocus()
    Call Grid_Recebe_Foco(objGridParadas)
End Sub

Private Sub GridParadas_EnterCell()
    If Not bDesabilitaCmdGridAux Then
        Call Grid_Entrada_Celula(objGridParadas, iAlterado)
    End If
End Sub

Private Sub GridParadas_LeaveCell()
    If Not bDesabilitaCmdGridAux Then
        Call Saida_Celula(objGridParadas)
    End If
End Sub

Private Sub GridParadas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridParadas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParadas, iAlterado)
    End If

End Sub

Private Sub GridParadas_RowColChange()
    If Not bDesabilitaCmdGridAux Then

        Call Grid_RowColChange(objGridParadas)
        
        Call Recolhe_Dados(iLinhaAnt)
        Call Mostra_Dados(GridParadas.Row)
                
    End If
    
    iLinhaAnt = GridParadas.Row
    LinhaAtual.Caption = CStr(GridParadas.Row)
    
End Sub

Private Sub GridParadas_Scroll()
    Call Grid_Scroll(objGridParadas)
End Sub

Private Sub GridParadas_KeyDown(KeyCode As Integer, Shift As Integer)
    
Dim lErro As Long
Dim iItemAtual As Integer
Dim iLinhasExistentesAnt As Integer
Dim vbMsgRes As VbMsgBoxResult
    
On Error GoTo Erro_GridParadas_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnt = objGridParadas.iLinhasExistentes
    iItemAtual = GridParadas.Row
    
    lErro = Remove_Linha(iItemAtual, KeyCode)
    If lErro <> SUCESSO Then gError 205223
    
    Call Grid_Trata_Tecla1(KeyCode, objGridParadas)

    Exit Sub

Erro_GridParadas_KeyDown:

    Select Case gErr

        Case 205223

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205224)

    End Select

    Exit Sub
    
End Sub

Private Sub GridParadas_LostFocus()
    Call Grid_Libera_Foco(objGridParadas)
End Sub

Private Sub ParadasSel_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ParadasSel_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParadas)
End Sub

Private Sub ParadasSel_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParadas)
End Sub

Private Sub ParadasSel_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParadas.objControle = ParadasSel
    lErro = Grid_Campo_Libera_Foco(objGridParadas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ParadasCliente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ParadasCliente_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParadas)
End Sub

Private Sub ParadasCliente_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParadas)
End Sub

Private Sub ParadasCliente_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParadas.objControle = ParadasCliente
    lErro = Grid_Campo_Libera_Foco(objGridParadas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ParadasEndereco_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ParadasEndereco_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParadas)
End Sub

Private Sub ParadasEndereco_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParadas)
End Sub

Private Sub ParadasEndereco_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParadas.objControle = ParadasEndereco
    lErro = Grid_Campo_Libera_Foco(objGridParadas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ParadasBairro_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ParadasBairro_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParadas)
End Sub

Private Sub ParadasBairro_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParadas)
End Sub

Private Sub ParadasBairro_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParadas.objControle = ParadasBairro
    lErro = Grid_Campo_Libera_Foco(objGridParadas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ParadasDistancia_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ParadasDistancia_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParadas)
End Sub

Private Sub ParadasDistancia_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParadas)
End Sub

Private Sub ParadasDistancia_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParadas.objControle = ParadasDistancia
    lErro = Grid_Campo_Libera_Foco(objGridParadas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ParadasOBS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ParadasOBS_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParadas)
End Sub

Private Sub ParadasOBS_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParadas)
End Sub

Private Sub ParadasOBS_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParadas.objControle = ParadasOBS
    lErro = Grid_Campo_Libera_Foco(objGridParadas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_ParadasCliente(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer

On Error GoTo Erro_Saida_Celula_ParadasCliente

    Set objGridInt.objControle = ParadasCliente

    If Len(Trim(ParadasCliente.ClipText)) > 0 Then
    
        If InStr(1, ParadasCliente.Text, "-") <> 0 Then ParadasCliente.Text = LCodigo_Extrai(ParadasCliente.Text)
    
        lErro = TP_Cliente_Le3(ParadasCliente, objcliente, iCodFilial)
        If lErro <> SUCESSO Then gError 205195
        
        objcliente.iFilialEmpresaLoja = iCodFilial
        
        lErro = Trata_InclusaoCliente(objcliente)
        If lErro <> SUCESSO Then gError 205205
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 205196

    Saida_Celula_ParadasCliente = SUCESSO

    Exit Function

Erro_Saida_Celula_ParadasCliente:

    Saida_Celula_ParadasCliente = gErr

    Select Case gErr

        Case 205195, 205196, 205205
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205197)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_ParadasFilial(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer

On Error GoTo Erro_Saida_Celula_ParadasFilial

    Set objGridInt.objControle = ParadasFilial

    If Len(Trim(ParadasFilial.ClipText)) = 0 Then gError 205206
        
    objcliente.lCodigo = LCodigo_Extrai(GridParadas.TextMatrix(GridParadas.Row, iGrid_ParadasCliente_Col))
    objcliente.iFilialEmpresaLoja = LCodigo_Extrai(ParadasFilial.Text)
    
    'Se alterou a filial
    If objcliente.iFilialEmpresaLoja <> Codigo_Extrai(GridParadas.TextMatrix(GridParadas.Row, iGrid_ParadasFilial_Col)) Then
        lErro = Trata_InclusaoCliente(objcliente)
        If lErro <> SUCESSO Then gError 205205
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 205196

    Saida_Celula_ParadasFilial = SUCESSO

    Exit Function

Erro_Saida_Celula_ParadasFilial:

    Saida_Celula_ParadasFilial = gErr

    Select Case gErr

        Case 205195, 205196, 205205
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 205206
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205197)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'GridVend
        If objGridInt.objGrid.Name = GridVend.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_Vendedor_Col

                    lErro = Saida_Celula_Vendedor(objGridInt)
                    If lErro <> SUCESSO Then gError 205198

            End Select
                    
        End If

        'GridParadas
        If objGridInt.objGrid.Name = GridParadas.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_ParadasSel_Col

                    lErro = Saida_Celula_Padrao(objGridInt, ParadasSel)
                    If lErro <> SUCESSO Then gError 205199

                Case iGrid_ParadasCliente_Col

                    lErro = Saida_Celula_ParadasCliente(objGridInt)
                    If lErro <> SUCESSO Then gError 205200

                Case iGrid_ParadasFilial_Col

                    lErro = Saida_Celula_ParadasFilial(objGridInt)
                    If lErro <> SUCESSO Then gError 205200

                Case iGrid_ParadasOBS_Col

                    lErro = Saida_Celula_Padrao(objGridInt, ParadasOBS)
                    If lErro <> SUCESSO Then gError 205201

            End Select
                    
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 205202

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 205198 To 205201

        Case 205202
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205203)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa o controle da coluna em questão
    Select Case objControl.Name
    
        Case ParadasCliente.Name
            If Len(Trim(GridParadas.TextMatrix(iLinha, iGrid_ParadasCliente_Col))) > 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case ParadasFilial.Name, ParadasOBS.Name, ParadasSel.Name
            If Len(Trim(GridParadas.TextMatrix(iLinha, iGrid_ParadasCliente_Col))) > 0 Then
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205204)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objcliente As ClassCliente
Dim bCancel As Boolean

On Error GoTo Erro_objEventoCliente_evSelecao

    Set objcliente = obj1

    'Preenche campo ParadasCliente
    If Me.ActiveControl Is ParadasCliente Then
        ParadasCliente.Text = objcliente.sNomeReduzido
    Else
        GridParadas.TextMatrix(GridParadas.Row, iGrid_ParadasCliente_Col) = objcliente.lCodigo & SEPARADOR & objcliente.sNomeReduzido
    
        lErro = Trata_InclusaoCliente(objcliente)
        If lErro <> SUCESSO Then gError 205205
    
    End If

    Me.Show

    Exit Sub
    
Erro_objEventoCliente_evSelecao:
    
    Select Case gErr
    
        Case 205205

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205206)
    
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoClientes_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection
Dim sFiltro As String

On Error GoTo Erro_BotaoClientes_Click

    If GridParadas.Row = 0 Then gError 205208

    If Me.ActiveControl Is ParadasCliente Then
        objcliente.lCodigo = LCodigo_Extrai(ParadasCliente.Text)
        objcliente.sNomeReduzido = ParadasCliente.Text
    Else
        If Me.ActiveControl Is ParadasFilial Then
            sFiltro = "ClienteCod = ?"
            colSelecao.Add LCodigo_Extrai(GridParadas.TextMatrix(GridParadas.Row, iGrid_ParadasCliente_Col))
        Else
            objcliente.lCodigo = LCodigo_Extrai(GridParadas.TextMatrix(GridParadas.Row, iGrid_ParadasCliente_Col))
        End If
    End If

    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente, sFiltro)

    Exit Sub
    
Erro_BotaoClientes_Click:
    
    Select Case gErr
    
        Case 205208
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205209)
    
    End Select
    
    Exit Sub

End Sub

Private Function Trata_InclusaoCliente(ByVal objcliente As ClassCliente) As Long

Dim lErro As Long
Dim objPonto As New ClassRotasPontos
Dim objEndereco As New ClassEndereco
Dim objFilial As New ClassFilialCliente

On Error GoTo Erro_Trata_InclusaoCliente

    ParadasCliente.Text = CStr(objcliente.lCodigo) & SEPARADOR & objcliente.sNomeReduzido

    Set objPonto.objcliente = objcliente
    If gobjRota.colPontos.Count > 0 Then Set objPonto.objPontoAnt = gobjRota.colPontos.Item(gobjRota.colPontos.Count)

    objFilial.lCodCliente = objcliente.lCodigo
    objFilial.iCodFilial = objcliente.iFilialEmpresaLoja
    If objFilial.iCodFilial = 0 Then
        objFilial.iCodFilial = FILIAL_MATRIZ
    End If

    lErro = CF("FilialCliente_Le", objFilial)
    If lErro <> SUCESSO And lErro <> 12567 Then gError 205211
    
    If lErro <> SUCESSO Then gError 205260
    
    objEndereco.lCodigo = objFilial.lEnderecoEntrega
    
    'Le endereço no BD
    lErro = CF("Endereco_le", objEndereco)
    If lErro <> SUCESSO Then gError 205211
    
    If Len(Trim(objEndereco.sEndereco)) = 0 Then
    
        objEndereco.lCodigo = objFilial.lEndereco
    
        lErro = CF("Endereco_le", objEndereco)
        If lErro <> SUCESSO Then gError 205211
    
    End If
    
    Set objPonto.objEndereco = objEndereco
    
    objPonto.iSeq = gobjRota.colPontos.Count + 1
    objPonto.lCliente = objcliente.lCodigo
    
    ParadasFilial.Text = CStr(objFilial.iCodFilial) & SEPARADOR & objFilial.sNome
    GridParadas.TextMatrix(GridParadas.Row, iGrid_ParadasFilial_Col) = ParadasFilial.Text
    GridParadas.TextMatrix(GridParadas.Row, iGrid_ParadasEndereco_Col) = objEndereco.sEndereco
    GridParadas.TextMatrix(GridParadas.Row, iGrid_ParadasBairro_Col) = objEndereco.sBairro
    GridParadas.TextMatrix(GridParadas.Row, iGrid_ParadasSel_Col) = MARCADO

    gobjRota.colPontos.Add objPonto
    
    'verifica se precisa preencher o grid com uma nova linha
    If objGridParadas.objGrid.Row - objGridParadas.objGrid.FixedRows = objGridParadas.iLinhasExistentes Then
        objGridParadas.iLinhasExistentes = objGridParadas.iLinhasExistentes + 1
    End If

    Call Grid_Refresh_Checkbox(objGridParadas)

    Trata_InclusaoCliente = SUCESSO

    Exit Function

Erro_Trata_InclusaoCliente:

    Trata_InclusaoCliente = gErr

    Select Case gErr
    
        Case 205211
        
        Case 205260 'ERRO_FILIALCLIENTE_NAO_ENCONTRADA
             Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, objFilial.iCodFilial)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205210)

    End Select

    Exit Function
    
End Function

Private Sub TrechoMeio_Change()
    iAlterado = MARCADO
End Sub

Private Sub TrechoDistancia_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TrechoDistancia_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TrechoDistancia_Validate

    'Veifica se TrechoDistancia está preenchida
    If Len(Trim(TrechoDistancia.Text)) <> 0 Then

       'Critica a TrechoDistancia
       lErro = Valor_Positivo_Critica(TrechoDistancia.Text)
       If lErro <> SUCESSO Then gError 205212
       
       TrechoDistancia.Text = Formata_Estoque(TrechoDistancia.Text)
        
    End If

    Exit Sub

Erro_TrechoDistancia_Validate:

    Cancel = True

    Select Case gErr

        Case 205212

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205213)

    End Select

    Exit Sub
    
End Sub

Private Sub TrechoTempo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TrechoTempo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TrechoTempo_Validate

    'Veifica se TrechoTempo está preenchida
    If Len(Trim(TrechoTempo.Text)) <> 0 Then

       'Critica a TrechoTempo
       lErro = Valor_Positivo_Critica(TrechoTempo.Text)
       If lErro <> SUCESSO Then gError 205214
       
       TrechoTempo.Text = Formata_Estoque(TrechoTempo.Text)
        
    End If

    Exit Sub

Erro_TrechoTempo_Validate:

    Cancel = True

    Select Case gErr

        Case 205214

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205215)

    End Select

    Exit Sub
    
End Sub

Private Function Mostra_Dados(ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim objPonto As ClassRotasPontos
Dim sEndereco As String
Dim objEnd As ClassEndereco
Dim objCamposGenericosValores As ClassCamposGenericosValores

On Error GoTo Erro_Mostra_Dados

    TrechoDistancia.Text = ""
    TrechoTempo.Text = ""
    TrechoMeio.ListIndex = -1
    TrechoEndAte.Caption = ""
    TrechoEndDe.Caption = ""
    FrameTrecho.Enabled = False
    FrameTrecho.Caption = "Trecho"
    DistAte.Caption = ""

    If iLinha <> 0 And gobjRota.colPontos.Count >= iLinha And Not bTrazendoDados Then
    
        If iLinha > 1 Then FrameTrecho.Enabled = True
    
        Set objPonto = gobjRota.colPontos.Item(iLinha)
        Set objEnd = objPonto.objEndereco
     
        If objPonto.lMeio <> 0 Then
            Call Combo_Seleciona_ItemData(TrechoMeio, objPonto.lMeio)
            For Each objCamposGenericosValores In gobjMeios.colCamposGenericosValores
                If objPonto.lMeio = objCamposGenericosValores.lCodValor Then
                    If IsNumeric(objCamposGenericosValores.sComplemento1) Then
                        DistAte.Caption = Formata_Estoque(objCamposGenericosValores.sComplemento1)
                    End If
                    Exit For
                End If
            Next
        End If
        
        If objPonto.dDistancia <> 0 Then TrechoDistancia.Text = Formata_Estoque(objPonto.dDistancia)
        If objPonto.dTempo <> 0 Then TrechoTempo.Text = Formata_Estoque(objPonto.dTempo)
        
        sEndereco = objEnd.sEndereco & " " & SEPARADOR & " " & objEnd.sBairro & " " & SEPARADOR & " " & objEnd.sCidade & " " & SEPARADOR & " " & objEnd.sSiglaEstado
        sEndereco = sEndereco & vbNewLine & "Tel1: " & objEnd.sTelefone1 & " Tel2: " & objEnd.sTelNumero2
        sEndereco = sEndereco & vbNewLine & "Ref.: " & objEnd.sReferencia
        
        TrechoEndAte.Caption = sEndereco
        
        If Not (objPonto.objPontoAnt Is Nothing) Then
        
            Set objEnd = objPonto.objPontoAnt.objEndereco

            sEndereco = objEnd.sEndereco & " " & SEPARADOR & " " & objEnd.sBairro & " " & SEPARADOR & " " & objEnd.sCidade & " " & SEPARADOR & " " & objEnd.sSiglaEstado
            sEndereco = sEndereco & vbNewLine & "Tel1: " & objEnd.sTelefone1 & " Tel2: " & objEnd.sTelNumero2
            sEndereco = sEndereco & vbNewLine & "Ref.: " & objEnd.sReferencia
            
            TrechoEndDe.Caption = sEndereco
            
            FrameTrecho.Caption = "Trecho: " & CStr(objPonto.objPontoAnt.objcliente.lCodigo) & " " & SEPARADOR & " " & objPonto.objPontoAnt.objcliente.sNomeReduzido & "\" & CStr(objPonto.objcliente.lCodigo) & " " & SEPARADOR & " " & objPonto.objcliente.sNomeReduzido
        
        End If

    End If
    
    Mostra_Dados = SUCESSO
    
    Exit Function

Erro_Mostra_Dados:

    Mostra_Dados = gErr

    Select Case gErr
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205216)

    End Select

    Exit Function

End Function

Private Function Recolhe_Dados(ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim objPonto As ClassRotasPontos
Dim iLinhaAux As Integer
Dim iColAux As Integer

On Error GoTo Erro_Recolhe_Dados

    If iLinha <> 0 And gobjRota.colPontos.Count >= iLinha And Not bTrazendoDados Then
    
        Set objPonto = gobjRota.colPontos.Item(iLinha)
        
        If Len(Trim(TrechoMeio.Text)) > 0 Then
            objPonto.lMeio = TrechoMeio.ItemData(TrechoMeio.ListIndex)
        Else
            objPonto.lMeio = 0
        End If
        objPonto.dDistancia = StrParaDbl(TrechoDistancia.Text)
        objPonto.dTempo = StrParaDbl(TrechoTempo.Text)
            
        If StrParaDbl(TrechoDistancia.Text) <> 0 Then
            GridParadas.TextMatrix(iLinha, iGrid_ParadasDistancia_Col) = Formata_Estoque(TrechoDistancia.Text)
        Else
            GridParadas.TextMatrix(iLinha, iGrid_ParadasDistancia_Col) = ""
        End If
            
        Call Soma_Coluna_Grid(objGridParadas, iGrid_ParadasDistancia_Col, DistTotal, False)

        iLinhaAux = GridParadas.Row
        iColAux = GridParadas.Col
        
        bDesabilitaCmdGridAux = True
        GridParadas.Row = iLinha
        GridParadas.Col = iGrid_ParadasDistancia_Col
        If objPonto.dDistancia > StrParaDbl(DistAte.Caption) And StrParaDbl(DistAte.Caption) > 0 Then
            GridParadas.CellForeColor = vbRed
        Else
            GridParadas.CellForeColor = vbBlack
        End If
        bDesabilitaCmdGridAux = False

        bDesabilitaCmdGridAux = True
        GridParadas.Row = iLinhaAux
        GridParadas.Col = iColAux
        bDesabilitaCmdGridAux = False
    End If
    
    Recolhe_Dados = SUCESSO
    
    Exit Function

Erro_Recolhe_Dados:

    bDesabilitaCmdGridAux = False

    Recolhe_Dados = gErr

    Select Case gErr
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205217)

    End Select

    Exit Function

End Function

Private Sub BotaoDesce_Click()
    Call Recolhe_Dados(iLinhaAnt)
    Call Troca_Dados_Posicao(GridParadas.Row, GridParadas.Row + 1)
End Sub

Private Sub BotaoFundo_Click()
    Call Recolhe_Dados(iLinhaAnt)
    Call Troca_Dados_Posicao(GridParadas.Row, objGridParadas.iLinhasExistentes)
End Sub

Private Sub BotaoSobe_Click()
    Call Recolhe_Dados(iLinhaAnt)
    Call Troca_Dados_Posicao(GridParadas.Row, GridParadas.Row - 1)
End Sub

Private Sub BotaoTopo_Click()
    Call Recolhe_Dados(iLinhaAnt)
    Call Troca_Dados_Posicao(GridParadas.Row, 1)
End Sub

Private Sub BotaoMudaLinha_Click()
    Call Recolhe_Dados(iLinhaAnt)
    Call Troca_Dados_Posicao(GridParadas.Row, StrParaInt(LinhaDesejada.Text))
End Sub

'Private Function Troca_Dados_Posicao(ByVal iLinha1 As Integer, ByVal iLinha2 As Integer) As Long
'
'Dim lErro As Long
'Dim objPontoAux As New ClassRotasPontos
'Dim objPonto2Aux As New ClassRotasPontos
'Dim objPonto As ClassRotasPontos
'Dim objPonto2 As ClassRotasPontos
'Dim bVaiPerderdadosTrecho As Boolean
'Dim vbMsg As VbMsgBoxResult
'Dim sNomeCli1 As String, sNomeCli2 As String, sNomeFil1 As String, sNomeFil2 As String
'
'On Error GoTo Erro_Troca_Dados_Posicao
'
'    If iLinha1 < 1 Or iLinha1 > gobjRota.ColPontos.Count Then gError 205219
'    If iLinha2 < 1 Or iLinha2 > gobjRota.ColPontos.Count Then gError 205220
'
'    bVaiPerderdadosTrecho = False
'
'    If iLinha1 < gobjRota.ColPontos.Count Then
'        Set objPonto = gobjRota.ColPontos.Item(iLinha1 + 1)
'        If objPonto.dDistancia <> 0 Or objPonto.dTempo <> 0 Or objPonto.lMeio <> 0 Then bVaiPerderdadosTrecho = True
'    End If
'
'    If iLinha2 < gobjRota.ColPontos.Count Then
'        Set objPonto2 = gobjRota.ColPontos.Item(iLinha2 + 1)
'        If objPonto2.dDistancia <> 0 Or objPonto2.dTempo <> 0 Or objPonto2.lMeio <> 0 Then bVaiPerderdadosTrecho = True
'    End If
'
'    Set objPonto = gobjRota.ColPontos.Item(iLinha1)
'    Set objPonto2 = gobjRota.ColPontos.Item(iLinha2)
'
'    If objPonto.dDistancia <> 0 Or objPonto.dTempo <> 0 Or objPonto.lMeio <> 0 Then bVaiPerderdadosTrecho = True
'    If objPonto2.dDistancia <> 0 Or objPonto2.dTempo <> 0 Or objPonto2.lMeio <> 0 Then bVaiPerderdadosTrecho = True
'
'    vbMsg = vbYes
'    If bVaiPerderdadosTrecho Then vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_TRECHO")
'    If vbMsg = vbNo Then gError 205261
'
'    'Joga o item para um obj Auxiliar
'    sNomeCli1 = GridParadas.TextMatrix(iLinha1, iGrid_ParadasCliente_Col)
'    sNomeFil1 = GridParadas.TextMatrix(iLinha1, iGrid_ParadasFilial_Col)
'    objPontoAux.lCliente = LCodigo_Extrai(sNomeCli1)
'    objPontoAux.iFilialCliente = Codigo_Extrai(sNomeFil1)
'    objPontoAux.iSeq = gobjRota.ColPontos.Item(iLinha1).iSeq
'    objPontoAux.sObservacao = GridParadas.TextMatrix(iLinha1, iGrid_ParadasOBS_Col)
'    objPontoAux.iSelecionado = StrParaInt(GridParadas.TextMatrix(iLinha1, iGrid_ParadasSel_Col))
'    Set objPontoAux.objEndereco = gobjRota.ColPontos.Item(iLinha1).objEndereco
'    Set objPontoAux.objcliente = gobjRota.ColPontos.Item(iLinha1).objcliente
'    If iLinha1 > 1 Then Set objPontoAux.objPontoAnt = gobjRota.ColPontos.Item(iLinha1).objPontoAnt
'
'    'Joga o item para dois obj Auxiliar
'    sNomeCli2 = GridParadas.TextMatrix(iLinha2, iGrid_ParadasCliente_Col)
'    sNomeFil2 = GridParadas.TextMatrix(iLinha2, iGrid_ParadasFilial_Col)
'    objPonto2Aux.lCliente = LCodigo_Extrai(sNomeCli2)
'    objPonto2Aux.iFilialCliente = Codigo_Extrai(sNomeFil2)
'    objPonto2Aux.iSeq = gobjRota.ColPontos.Item(iLinha2).iSeq
'    objPonto2Aux.sObservacao = GridParadas.TextMatrix(iLinha2, iGrid_ParadasOBS_Col)
'    objPonto2Aux.iSelecionado = StrParaInt(GridParadas.TextMatrix(iLinha2, iGrid_ParadasSel_Col))
'    Set objPonto2Aux.objEndereco = gobjRota.ColPontos.Item(iLinha2).objEndereco
'    Set objPonto2Aux.objcliente = gobjRota.ColPontos.Item(iLinha2).objcliente
'    If iLinha2 > 1 Then Set objPonto2Aux.objPontoAnt = gobjRota.ColPontos.Item(iLinha2).objPontoAnt
'
'    'Troca as posições de Ambos
'
'    objPonto2.lCliente = objPontoAux.lCliente
'    objPonto2.iFilialCliente = objPontoAux.iFilialCliente
'    objPonto2.iSeq = objPontoAux.iSeq
'    objPonto2.sObservacao = objPontoAux.sObservacao
'    objPonto2.iSelecionado = objPontoAux.iSelecionado
'    Set objPonto2.objEndereco = objPontoAux.objEndereco
'    Set objPonto2.objcliente = objPontoAux.objcliente
'    If Not (objPontoAux.objPontoAnt Is Nothing) Then Set objPonto2.objPontoAnt = objPontoAux.objPontoAnt
'
'    objPonto.lCliente = objPonto2Aux.lCliente
'    objPonto.iFilialCliente = objPonto2Aux.iFilialCliente
'    objPonto.iSeq = objPonto2Aux.iSeq
'    objPonto.sObservacao = objPonto2Aux.sObservacao
'    objPonto.iSelecionado = objPonto2Aux.iSelecionado
'    Set objPonto.objEndereco = objPonto2Aux.objEndereco
'    Set objPonto.objcliente = objPonto2Aux.objcliente
'    If Not (objPonto2Aux.objPontoAnt Is Nothing) Then Set objPonto.objPontoAnt = objPonto2Aux.objPontoAnt
'
'    'Coloca as informações no Grid
'    GridParadas.TextMatrix(iLinha1, iGrid_ParadasEndereco_Col) = objPonto.objEndereco.sEndereco
'    GridParadas.TextMatrix(iLinha1, iGrid_ParadasBairro_Col) = objPonto.objEndereco.sBairro
'    GridParadas.TextMatrix(iLinha1, iGrid_ParadasOBS_Col) = objPonto.sObservacao
'    GridParadas.TextMatrix(iLinha1, iGrid_ParadasDistancia_Col) = ""
'    GridParadas.TextMatrix(iLinha1, iGrid_ParadasSel_Col) = objPonto.iSelecionado
'    GridParadas.TextMatrix(iLinha1, iGrid_ParadasCliente_Col) = sNomeCli2
'    GridParadas.TextMatrix(iLinha1, iGrid_ParadasFilial_Col) = sNomeFil2
'
'    GridParadas.TextMatrix(iLinha2, iGrid_ParadasEndereco_Col) = objPonto2.objEndereco.sEndereco
'    GridParadas.TextMatrix(iLinha2, iGrid_ParadasBairro_Col) = objPonto2.objEndereco.sBairro
'    GridParadas.TextMatrix(iLinha2, iGrid_ParadasOBS_Col) = objPonto2.sObservacao
'    GridParadas.TextMatrix(iLinha2, iGrid_ParadasDistancia_Col) = ""
'    GridParadas.TextMatrix(iLinha2, iGrid_ParadasSel_Col) = objPonto2.iSelecionado
'    GridParadas.TextMatrix(iLinha2, iGrid_ParadasCliente_Col) = sNomeCli1
'    GridParadas.TextMatrix(iLinha2, iGrid_ParadasFilial_Col) = sNomeFil1
'
'    Call Limpa_Trecho(iLinha1 + 1)
'    Call Limpa_Trecho(iLinha2 + 1)
'
'    Call Mostra_Dados(GridParadas.Row)
'
'    Troca_Dados_Posicao = SUCESSO
'
'    Exit Function
'
'Erro_Troca_Dados_Posicao:
'
'    Troca_Dados_Posicao = gErr
'
'    Select Case gErr
'
'        Case 205219
'
'        Case 205220
'             Call Rotina_Erro(vbOKOnly, "ERRO_MUDANCA_LINHA_INVALIDA", gErr, iLinha2, 1, gobjRota.ColPontos.Count)
'
'        Case 205261
'
'        Case Else
'             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205218)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Troca_Dados_Posicao(ByVal iLinha1 As Integer, ByVal iLinha2 As Integer) As Long

Dim lErro As Long
Dim objPonto As ClassRotasPontos
Dim bVaiPerderdadosTrecho As Boolean
Dim vbMsg As VbMsgBoxResult
Dim iPonto As Integer
Dim colPontos As New Collection, colCampos As New Collection

On Error GoTo Erro_Troca_Dados_Posicao

    If iLinha1 < 1 Or iLinha1 > gobjRota.colPontos.Count Then gError 205219
    If iLinha2 < 1 Or iLinha2 > gobjRota.colPontos.Count Then gError 205220

    bVaiPerderdadosTrecho = False

    'Verifica se o ponto posterior atual tem dados do trecho preenchidos
    If iLinha1 < gobjRota.colPontos.Count Then
        Set objPonto = gobjRota.colPontos.Item(iLinha1 + 1)
        If objPonto.dDistancia <> 0 Or objPonto.dTempo <> 0 Or objPonto.lMeio <> 0 Then bVaiPerderdadosTrecho = True
    End If
    
    'Se vai dminuir na sequência
    If iLinha1 > iLinha2 Then
        'Verifica se o ponto para onde vai tem dados do trecho preenchidos
        Set objPonto = gobjRota.colPontos.Item(iLinha2)
        If objPonto.dDistancia <> 0 Or objPonto.dTempo <> 0 Or objPonto.lMeio <> 0 Then bVaiPerderdadosTrecho = True
    Else
        'Verifica se o ponto seguinte para onde vai tem dados do trecho preenchidos
        If iLinha2 < gobjRota.colPontos.Count Then
            Set objPonto = gobjRota.colPontos.Item(iLinha2 + 1)
            If objPonto.dDistancia <> 0 Or objPonto.dTempo <> 0 Or objPonto.lMeio <> 0 Then bVaiPerderdadosTrecho = True
        End If
    End If
    
    'Verifica se o próprio tem dados do trecho preenchidos
    Set objPonto = gobjRota.colPontos.Item(iLinha1)
    If objPonto.dDistancia <> 0 Or objPonto.dTempo <> 0 Or objPonto.lMeio <> 0 Then bVaiPerderdadosTrecho = True
    
    vbMsg = vbYes
    'Se for peder dados pede confirmação
    If bVaiPerderdadosTrecho Then vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_TRECHO")
    If vbMsg = vbNo Then gError 205261
    
    If iLinha1 > iLinha2 Then
        iPonto = 0
        For Each objPonto In gobjRota.colPontos
            iPonto = iPonto + 1
            If iPonto < iLinha1 And iPonto >= iLinha2 Then
                objPonto.iSeq = objPonto.iSeq + 1
            End If
            If iPonto = iLinha1 Then objPonto.iSeq = iLinha2
        Next
    Else
        iPonto = 0
        For Each objPonto In gobjRota.colPontos
            iPonto = iPonto + 1
            If iPonto > iLinha1 And iPonto <= iLinha2 Then
                objPonto.iSeq = objPonto.iSeq - 1
            End If
            If iPonto = iLinha1 Then objPonto.iSeq = iLinha2
        Next
    End If
    
    colCampos.Add "iSeq"
    Call Ordena_Colecao(gobjRota.colPontos, colPontos, colCampos)
    
    For iPonto = gobjRota.colPontos.Count To 1 Step -1
        gobjRota.colPontos.Remove iPonto
    Next
    For Each objPonto In colPontos
        gobjRota.colPontos.Add objPonto
    Next
        
    lErro = Traz_Pontos_Tela(gobjRota)
    If lErro <> SUCESSO Then gError 205261
    
    If iLinha1 > iLinha2 Then
        Call Limpa_Trecho(iLinha1 + 1)
        Call Limpa_Trecho(iLinha2)
        Call Limpa_Trecho(iLinha2 + 1)
    Else
        Call Limpa_Trecho(iLinha1)
        Call Limpa_Trecho(iLinha2)
        Call Limpa_Trecho(iLinha2 + 1)
    End If
    
    bDesabilitaCmdGridAux = True
    GridParadas.Row = iLinha2
    bDesabilitaCmdGridAux = False
    
    Call Mostra_Dados(GridParadas.Row)
    
    Troca_Dados_Posicao = SUCESSO
    
    Exit Function

Erro_Troca_Dados_Posicao:

    Troca_Dados_Posicao = gErr

    Select Case gErr
    
        Case 205219
        
        Case 205220
             Call Rotina_Erro(vbOKOnly, "ERRO_MUDANCA_LINHA_INVALIDA", gErr, iLinha2, 1, gobjRota.colPontos.Count)
        
        Case 205261
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205218)

    End Select

    Exit Function

End Function

Public Function Remove_Linha(ByVal iLinha As Integer, ByVal iKeyCode As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Remove_Linha

    If iKeyCode = vbKeyDelete Then
        gobjRota.colPontos.Remove (iLinha)
        Call Limpa_Trecho(iLinha)
    End If
    
    Remove_Linha = SUCESSO
        
    Exit Function

Erro_Remove_Linha:

    Remove_Linha = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 205221)

    End Select

    Exit Function

End Function

Public Function Limpa_Trecho(ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim objPonto As ClassRotasPontos

On Error GoTo Erro_Limpa_Trecho

    If iLinha <= gobjRota.colPontos.Count And iLinha <> 0 Then
    
        Set objPonto = gobjRota.colPontos.Item(iLinha)
    
        If iLinha > 1 Then
            Set objPonto.objPontoAnt = gobjRota.colPontos.Item(iLinha - 1)
        Else
            Set objPonto.objPontoAnt = Nothing
        End If
        objPonto.dDistancia = 0
        objPonto.dTempo = 0
        objPonto.lMeio = 0
        
        GridParadas.TextMatrix(iLinha, iGrid_ParadasDistancia_Col) = ""
    
    End If
    
    Call Soma_Coluna_Grid(objGridParadas, iGrid_ParadasDistancia_Col, DistTotal, False)
    
    Limpa_Trecho = SUCESSO
        
    Exit Function

Erro_Limpa_Trecho:

    Limpa_Trecho = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 205225)

    End Select

    Exit Function

End Function

Public Sub BotaoVendedores_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoVendedores_Click
    
    If GridVend.Row = 0 Then gError 205226
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Me.ActiveControl Is Vendedor Then
        objVendedor.iCodigo = Codigo_Extrai(Vendedor.Text)
        objVendedor.sNomeReduzido = Vendedor.Text
    Else
        objVendedor.iCodigo = Codigo_Extrai(GridVend.TextMatrix(GridVend.Row, iGrid_Vendedor_Col))
    End If
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

    Exit Sub

Erro_BotaoVendedores_Click:

    Select Case gErr
        
        Case 205226
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205227)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim objEndereco As New ClassEndereco
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_objEventoVendedor_evSelecao

    Set objVendedor = obj1
        
    If Me.ActiveControl Is Vendedor Then
        Vendedor.Text = objVendedor.iCodigo & SEPARADOR & objVendedor.sNomeReduzido
    Else
        
        Vendedor.Text = CStr(objVendedor.iCodigo)
        
        'Verifica se Vendedor existe
        lErro = TP_Vendedor_Le2(Vendedor, objVendedor)
        If lErro <> SUCESSO And lErro <> 25018 And lErro <> 25020 Then gError 205228

        If lErro = 25018 Then gError 205229
        If lErro = 25020 Then gError 205230
    
        Call TP_Vendedor_Le2(Vendedor, objVendedor)
        GridVend.TextMatrix(GridVend.Row, iGrid_Vendedor_Col) = Vendedor.Text
    
        objEndereco.lCodigo = objVendedor.lEndereco
        
        'Le endereço no BD
        lErro = CF("Endereco_le", objEndereco)
        If lErro <> SUCESSO Then gError 205231
    
        GridVend.TextMatrix(GridVend.Row, iGrid_VendTelefone1_Col) = objEndereco.sTelefone1
        GridVend.TextMatrix(GridVend.Row, iGrid_VendTelefone2_Col) = objEndereco.sTelefone2

        If (GridVend.Row - GridVend.FixedRows) = objGridVend.iLinhasExistentes Then
            objGridVend.iLinhasExistentes = objGridVend.iLinhasExistentes + 1
        End If
    
    End If
    
    Me.Show

    Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case gErr
        
        Case 205228
            
        Case 205229 'Não encontrou nome reduzido de vendedor no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR")
            If vbMsgRes = vbYes Then
                'Preenche objVendedor com nome reduzido
                objVendedor.sNomeReduzido = Vendedor.Text

                'Chama a tela de Vendedores
                Call Chama_Tela("Vendedores", objVendedor)
            End If

        Case 205230 'Não encontrou codigo do vendedor no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR")
            If vbMsgRes = vbYes Then
                'Prenche objVendedor com codigo
                objVendedor.iCodigo = Codigo_Extrai(Vendedor.Text)

                'Chama a tela de Vendedores
                Call Chama_Tela("Vendedores", objVendedor)
            End If
            
        Case 205231
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205232)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoQuebrar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Quebra_Rota
    If lErro <> SUCESSO Then gError 205233

    'Limpa Tela
    Call Limpa_Tela_Rotas

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 205233

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205234)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoTransf_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Transf_Rota
    If lErro <> SUCESSO Then gError 205235

    'Limpa Tela
    Call Limpa_Tela_Rotas

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 205235

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205236)

    End Select

    Exit Sub
    
End Sub

Function Quebra_Rota() As Long

Dim lErro As Long
Dim objRota1 As New ClassRotas
Dim objRota2 As New ClassRotas
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Quebra_Rota

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Text)) = 0 Then gError 205237
    If Len(Trim(Chave.Text)) = 0 Then gError 205238
    If Len(Trim(CodigoTransf.Text)) = 0 Then gError 205239
    If Len(Trim(ChaveTransf.Text)) = 0 Then gError 205240

    'Preenche o objRota
    lErro = Move_Tela_Memoria(objRota1, objRota2, True)
    If lErro <> SUCESSO Then gError 205241
    
    'Pergunta ao usuário se confirma a quebra
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_QUEBRA_ROTAS", objRota1.sCodigo, objRota1.lChave, objRota2.sCodigo, objRota2.lChave)
    If vbMsgRes = vbNo Then gError 205249
    
    If objRota1.colPontos.Count = 0 Then gError 205242
    If objRota2.colPontos.Count = 0 Then gError 205243

    'Grava o/a Rotas no Banco de Dados
    lErro = CF("Rotas_Quebra", objRota1, objRota2)
    If lErro <> SUCESSO Then gError 205244

    GL_objMDIForm.MousePointer = vbDefault

    Quebra_Rota = SUCESSO

    Exit Function

Erro_Quebra_Rota:

    Quebra_Rota = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 205237
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_ROTA_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 205238
            Call Rotina_Erro(vbOKOnly, "ERRO_CHAVE_ROTA_NAO_PREENCHIDO", gErr)
            Chave.SetFocus
            
        Case 205239
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGOTRANSF_ROTA_NAO_PREENCHIDO", gErr)
            CodigoTransf.SetFocus

        Case 205240
            Call Rotina_Erro(vbOKOnly, "ERRO_CHAVETRANSF_ROTA_NAO_PREENCHIDO", gErr)
            ChaveTransf.SetFocus
            
        Case 205241, 205244, 205249
        
        Case 205242
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTAS_SEM_PARADAS'", gErr, objRota1.sCodigo, objRota1.lChave)

        Case 205243
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTAS_SEM_PARADAS'", gErr, objRota2.sCodigo, objRota2.lChave)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205245)

    End Select

    Exit Function

End Function

Function Transf_Rota() As Long

Dim lErro As Long
Dim objRota1 As New ClassRotas
Dim objRota2 As New ClassRotas
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Transf_Rota

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Text)) = 0 Then gError 205237
    If Len(Trim(Chave.Text)) = 0 Then gError 205238
    If Len(Trim(CodigoTransf.Text)) = 0 Then gError 205239
    If Len(Trim(ChaveTransf.Text)) = 0 Then gError 205240

    'Preenche o objRota
    lErro = Move_Tela_Memoria(objRota1, objRota2, True)
    If lErro <> SUCESSO Then gError 205241
    
    'Pergunta ao usuário se confirma a quebra
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_TRANSF_ROTAS", objRota1.sCodigo, objRota1.lChave, objRota2.sCodigo, objRota2.lChave)
    If vbMsgRes = vbNo Then gError 205249
    
    If objRota1.colPontos.Count = 0 Then gError 205242
    If objRota2.colPontos.Count = 0 Then gError 205243

    'Grava o/a Rotas no Banco de Dados
    lErro = CF("Rotas_Transf", objRota1, objRota2)
    If lErro <> SUCESSO Then gError 205244

    GL_objMDIForm.MousePointer = vbDefault

    Transf_Rota = SUCESSO

    Exit Function

Erro_Transf_Rota:

    Transf_Rota = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 205237
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_ROTA_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 205238
            Call Rotina_Erro(vbOKOnly, "ERRO_CHAVE_ROTA_NAO_PREENCHIDO", gErr)
            Chave.SetFocus
            
        Case 205239
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGOTRANSF_ROTA_NAO_PREENCHIDO", gErr)
            CodigoTransf.SetFocus

        Case 205240
            Call Rotina_Erro(vbOKOnly, "ERRO_CHAVETRANSF_ROTA_NAO_PREENCHIDO", gErr)
            ChaveTransf.SetFocus
            
        Case 205241, 205244, 205249
        
        Case 205242
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTAS_SEM_PARADAS'", gErr, objRota1.sCodigo, objRota1.lChave)

        Case 205243
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTAS_SEM_PARADAS'", gErr, objRota2.sCodigo, objRota2.lChave)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205245)

    End Select

    Exit Function

End Function

Private Sub LabelCodigoTransf_Click()

Dim lErro As Long
Dim objRota As New ClassRotas
Dim colSelecao As New Collection
Dim sFiltro As String

On Error GoTo Erro_LabelCodigoTransf_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(CodigoTransf.Text)) <> 0 Then

        objRota.sCodigo = CodigoTransf.Text

    End If
    
    If Len(Trim(ChaveTransf.Text)) > 0 Then
        sFiltro = "ChaveCod = ?"
        colSelecao.Add ChaveTransf.ItemData(ChaveTransf.ListIndex)
    End If

    Call Chama_Tela("RotasLista", colSelecao, objRota, objEventoRotaTransf, sFiltro)

    Exit Sub

Erro_LabelCodigoTransf_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205188)

    End Select

    Exit Sub

End Sub

Private Sub objEventoRotaTransf_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRota As ClassRotas

On Error GoTo Erro_objEventoRotaTransf_evSelecao

    Set objRota = obj1

    If objRota.sCodigo <> "" Then
        CodigoTransf.Text = objRota.sCodigo
    End If

    If objRota.lChave <> 0 Then
        Call Combo_Seleciona_ItemData(ChaveTransf, objRota.lChave)
    End If

    Me.Show

    Exit Sub

Erro_objEventoRotaTransf_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205187)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim dNumProx As Double

On Error GoTo Erro_BotaoProxNum_Click

    If Len(Trim(Chave.Text)) = 0 Then gError 205246

    'Gera número automático.
    lErro = CF("Rotas_Automatico", Chave.ItemData(Chave.ListIndex), dNumProx)
    If lErro <> SUCESSO Then gError 205247
    
    Codigo.Text = CStr(dNumProx)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 205246
            Call Rotina_Erro(vbOKOnly, "ERRO_CHAVE_ROTA_NAO_PREENCHIDO", gErr)
            Chave.SetFocus
            
        Case 205247

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205248)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoProxNumTransf_Click()

Dim lErro As Long
Dim dNumProx As Double

On Error GoTo Erro_BotaoProxNumTransf_Click

    If Len(Trim(ChaveTransf.Text)) = 0 Then gError 205246

    'Gera número automático.
    lErro = CF("Rotas_Automatico", ChaveTransf.ItemData(ChaveTransf.ListIndex), dNumProx)
    If lErro <> SUCESSO Then gError 205247
    
    CodigoTransf.Text = CStr(dNumProx)
    
    Exit Sub

Erro_BotaoProxNumTransf_Click:

    Select Case gErr

        Case 205246
            Call Rotina_Erro(vbOKOnly, "ERRO_CHAVETRANSF_ROTA_NAO_PREENCHIDO", gErr)
            ChaveTransf.SetFocus
            
        Case 205247

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205248)

    End Select
    
    Exit Sub
    
End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is ParadasCliente Then
            Call BotaoClientes_Click
        ElseIf Me.ActiveControl Is ParadasFilial Then
            Call BotaoClientes_Click
        ElseIf Me.ActiveControl Is CodigoTransf Then
            Call LabelCodigoTransf_Click
        ElseIf Me.ActiveControl Is Vendedor Then
            Call BotaoVendedores_Click
        End If
    
    End If

End Sub

Private Sub LinhaDesejada_GotFocus()
    Call MaskEdBox_TrataGotFocus(LinhaDesejada, iAlterado)
End Sub

Private Sub BotaoDesmarcarTodos_Click()
    Call Grid_Marca_Desmarca(objGridParadas, iGrid_ParadasSel_Col, DESMARCADO)
End Sub

Private Sub BotaoMarcarTodos_Click()
    Call Grid_Marca_Desmarca(objGridParadas, iGrid_ParadasSel_Col, MARCADO)
End Sub

Private Sub TrechoMeio_Click()
Dim objCamposGenericosValores As ClassCamposGenericosValores
    If TrechoMeio.ListIndex <> -1 Then
        For Each objCamposGenericosValores In gobjMeios.colCamposGenericosValores
            If TrechoMeio.ItemData(TrechoMeio.ListIndex) = objCamposGenericosValores.lCodValor Then
                If IsNumeric(objCamposGenericosValores.sComplemento1) Then
                    DistAte.Caption = Formata_Estoque(objCamposGenericosValores.sComplemento1)
                End If
                Exit For
            End If
        Next
    Else
        DistAte.Caption = ""
    End If
End Sub

Private Sub BotaoMapa_Click()

Dim lErro As Long
Dim lNumIntRel As Long
Dim sDiretorio As String
Dim lRetorno As Long
Dim objRotaPontos As ClassRotasPontos
Dim iIndice As Integer

On Error GoTo Erro_BotaoMapa_Click

    GL_objMDIForm.MousePointer = vbHourglass

    Call Recolhe_Dados(iLinhaAnt)
    
    If gobjFAT.iPossuiIntMapLink = DESMARCADO Then gError 205552
    
    For iIndice = 1 To objGridParadas.iLinhasExistentes
        Set objRotaPontos = gobjRota.colPontos.Item(iIndice)
        objRotaPontos.iSelecionado = StrParaInt(GridParadas.TextMatrix(iIndice, iGrid_ParadasSel_Col))
    Next
    
    lErro = CF("Rota_Exibe_Mapa_Prepara", gobjRota, lNumIntRel)
    If lErro <> SUCESSO Then gError 205553

    sDiretorio = String(255, 0)
    lRetorno = GetPrivateProfileString("Forprint", "DirBin", "c:\sge\programa\", sDiretorio, 255, NOME_ARQUIVO_ADM)
    sDiretorio = left(sDiretorio, lRetorno)
    
    lErro = WinExec(sDiretorio & "rota.exe 1 " & CStr(glEmpresa) & " " & CStr(lNumIntRel) & " 0 " & "Rota_" & CStr(Codigo.Text), SW_NORMAL)

    glNumIntRel = lNumIntRel
    gsDiretorio = sDiretorio
    
    Timer2.Enabled = True
    giContadorTempo = 0

    Exit Sub

Erro_BotaoMapa_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 205552
            Call Rotina_Aviso(vbOKOnly, "AVISO_FUNC_TERCEITOS_SEM_CONFIG")

        Case 205553

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205254)

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

    Call Recolhe_Dados(iLinhaAnt)
    
    If gobjFAT.iPossuiIntMapLink = DESMARCADO Then gError 205555
    
    lErro = CF("Rota_Seq_Mapa_Prepara", gobjRota, lNumIntRel)
    If lErro <> SUCESSO Then gError 205556

    sDiretorio = String(255, 0)
    lRetorno = GetPrivateProfileString("Forprint", "DirBin", "c:\sge\programa\", sDiretorio, 255, NOME_ARQUIVO_ADM)
    sDiretorio = left(sDiretorio, lRetorno)
    
    glNumIntRel = lNumIntRel
    gsDiretorio = sDiretorio

    giTentativa = 1
    lErro = WinExec(sDiretorio & "rota.exe 2 " & CStr(glEmpresa) & " " & CStr(lNumIntRel) & " 1 " & "Rota_" & CStr(Codigo.Text), SW_NORMAL)

    Timer1.Enabled = True
    giContadorTempo = 0

    Exit Sub

Erro_BotaoOrdAuto_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 205555
            Call Rotina_Aviso(vbOKOnly, "AVISO_FUNC_TERCEITOS_SEM_CONFIG")

        Case 205556

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205557)

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
        If vbResult = vbNo Then gError 205559
        GL_objMDIForm.MousePointer = vbHourglass
        giContadorTempo = 0
    End If
    
    lErro = CF("MapaRota1_Verifica_Retorno", glNumIntRel, sRetMsg)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then
    
        giTentativa = giTentativa + 1
        If giTentativa > NUM_TENTATIVAS Then gError 205558
    
        lErro2 = CF("Rota_Seq_Mapa_Prepara", gobjRota, lNumIntRel)
        If lErro2 <> SUCESSO Then gError 205559
        
        glNumIntRel = lNumIntRel
    
        Call WinExec(gsDiretorio & "rota.exe 2 " & CStr(glEmpresa) & " " & CStr(glNumIntRel) & " 1 " & "Rota_" & CStr(Codigo.Text), SW_NORMAL)
             
    End If
    
    If lErro = SUCESSO Then
        Timer1.Enabled = False
        giContadorTempo = 0
        
        lErro = CF("Rota_Seq_Mapa_Obtem", gobjRota, glNumIntRel)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 205560
        
        lErro = Traz_Pontos_Tela(gobjRota)
        If lErro <> SUCESSO Then gError 205561
        
        Call Mostra_Dados(GridParadas.Row)
        
        GL_objMDIForm.MousePointer = vbDefault
        
    End If

    Exit Sub

Erro_Timer1_Timer:

    Timer1.Enabled = False
    giContadorTempo = 0
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 205558
            Call Rotina_Erro(vbOKOnly, sRetMsg, gErr)
    
        Case 205559 To 205561

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205563)

    End Select
    
    Exit Sub
    
End Sub

Private Sub Timer2_Timer()

Const TEMPO_MAX_ESPERA = 20
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
        If vbResult = vbNo Then gError 205558
        GL_objMDIForm.MousePointer = vbHourglass
        giContadorTempo = 0
    End If
    
    lErro = CF("MapaRota_Verifica_Retorno", glNumIntRel, sRetMsg)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then
    
        giTentativa = giTentativa + 1
        If giTentativa > NUM_TENTATIVAS Then gError 205559
    
        lErro2 = CF("Rota_Exibe_Mapa_Prepara", gobjRota, lNumIntRel)
        If lErro2 <> SUCESSO Then gError 205558
        
        glNumIntRel = lNumIntRel
    
        Call WinExec(gsDiretorio & "rota.exe 1 " & CStr(glEmpresa) & " " & CStr(glNumIntRel) & " 0 " & "Rota_" & CStr(Codigo.Text), SW_NORMAL)
       
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
      
        Case 205558

        Case 205559
            Call Rotina_Erro(vbOKOnly, sRetMsg, gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205563)

    End Select
    
    Exit Sub
    
End Sub
