VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl BorderoCheque 
   ClientHeight    =   5925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9075
   ScaleHeight     =   5925
   ScaleWidth      =   9075
   Begin VB.Frame FrameCheque 
      BorderStyle     =   0  'None
      Caption         =   "Cheques Descritos"
      Height          =   2955
      Index           =   2
      Left            =   120
      TabIndex        =   43
      Top             =   2055
      Visible         =   0   'False
      Width           =   8850
      Begin VB.CheckBox SelecionadoN 
         BackColor       =   &H80000005&
         Height          =   195
         Left            =   2670
         TabIndex        =   48
         Top             =   510
         Width           =   1200
      End
      Begin MSMask.MaskEdBox DataDepositoChequeN 
         Height          =   300
         Left            =   3870
         TabIndex        =   49
         Top             =   435
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorN 
         Height          =   300
         Left            =   4875
         TabIndex        =   50
         Top             =   420
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin VB.CommandButton BotaoEditarN 
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
         Height          =   450
         Left            =   6705
         TabIndex        =   46
         Top             =   975
         Width           =   990
      End
      Begin VB.CommandButton BotaoMarcarTodosN 
         Caption         =   "Marcar Todos"
         Height          =   585
         Left            =   2865
         Picture         =   "BorderoCheque.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   2310
         Width           =   1440
      End
      Begin VB.CommandButton BotaoDesmarcarTodosN 
         Caption         =   "Desmarcar Todos"
         Height          =   585
         Left            =   4455
         Picture         =   "BorderoCheque.ctx":101A
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   2310
         Width           =   1440
      End
      Begin MSFlexGridLib.MSFlexGrid GridChequeN 
         Height          =   2130
         Left            =   2250
         TabIndex        =   47
         Top             =   120
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   3757
         _Version        =   393216
         Rows            =   5
         Cols            =   5
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame FrameCheque 
      BorderStyle     =   0  'None
      Caption         =   "Cheques Descritos"
      Height          =   2955
      Index           =   1
      Left            =   135
      TabIndex        =   31
      Top             =   2055
      Width           =   8850
      Begin VB.CheckBox Selecionado 
         BackColor       =   &H80000005&
         Height          =   195
         Left            =   210
         TabIndex        =   35
         Top             =   810
         Width           =   570
      End
      Begin MSMask.MaskEdBox CPFCGC 
         Height          =   300
         Left            =   6870
         TabIndex        =   36
         Top             =   765
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataDepositoCheque 
         Height          =   300
         Left            =   4665
         TabIndex        =   37
         Top             =   780
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Conta 
         Height          =   300
         Left            =   2940
         TabIndex        =   38
         Top             =   780
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   300
         Left            =   3810
         TabIndex        =   39
         Top             =   780
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Agencia 
         Height          =   300
         Left            =   2085
         TabIndex        =   40
         Top             =   795
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Banco 
         Height          =   300
         Left            =   1215
         TabIndex        =   41
         Top             =   795
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   300
         Left            =   5655
         TabIndex        =   42
         Top             =   765
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   585
         Left            =   2850
         Picture         =   "BorderoCheque.ctx":21FC
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2310
         Width           =   1440
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   585
         Left            =   4455
         Picture         =   "BorderoCheque.ctx":3216
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2310
         Width           =   1440
      End
      Begin MSFlexGridLib.MSFlexGrid GridCheques 
         Height          =   2130
         Left            =   -15
         TabIndex        =   34
         Top             =   90
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   3757
         _Version        =   393216
         Rows            =   5
         Cols            =   5
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton TestaLog 
      Caption         =   "LOG"
      Height          =   225
      Left            =   3210
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   1740
      Visible         =   0   'False
      Width           =   1965
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3315
      Left            =   60
      TabIndex        =   30
      Top             =   1725
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   5847
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Detalhados"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Não Detalhados"
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
   Begin VB.Frame FrameDestino 
      Caption         =   "Destino"
      Height          =   975
      Left            =   60
      TabIndex        =   25
      Top             =   660
      Width           =   5250
      Begin VB.OptionButton OptionBackoffice 
         Caption         =   "BackOffice"
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
         Left            =   405
         TabIndex        =   29
         Top             =   615
         Width           =   1650
      End
      Begin VB.OptionButton OptionConta 
         Caption         =   "Conta Corrente"
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
         Left            =   420
         TabIndex        =   28
         Top             =   255
         Value           =   -1  'True
         Width           =   1650
      End
      Begin VB.ComboBox ContaCorrente 
         Height          =   315
         Left            =   2790
         TabIndex        =   26
         Top             =   225
         Width           =   2190
      End
      Begin VB.Label LabelContaCorrente 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Conta:"
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
         Left            =   2160
         TabIndex        =   27
         Top             =   285
         Width           =   570
      End
   End
   Begin VB.Frame FrameBomPara 
      Caption         =   "Bom para"
      Height          =   1005
      Left            =   5415
      TabIndex        =   12
      Top             =   645
      Width           =   3585
      Begin VB.CommandButton BotaoTrazer 
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
         Height          =   450
         Left            =   2355
         TabIndex        =   22
         Top             =   345
         Width           =   990
      End
      Begin MSComCtl2.UpDown UpDownDe 
         Height          =   315
         Left            =   1860
         TabIndex        =   15
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDownAte 
         Height          =   315
         Left            =   1860
         TabIndex        =   16
         Top             =   585
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataBomParaDe 
         Height          =   300
         Left            =   870
         TabIndex        =   23
         Top             =   195
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataBomParaAte 
         Height          =   300
         Left            =   885
         TabIndex        =   24
         Top             =   585
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Até :"
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
         Left            =   390
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   14
         Top             =   645
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   " De :"
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
         Left            =   405
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   13
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2445
      Picture         =   "BorderoCheque.ctx":43F8
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Numeração Automática"
      Top             =   270
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6390
      ScaleHeight     =   495
      ScaleWidth      =   2520
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Width           =   2580
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1035
         Picture         =   "BorderoCheque.ctx":44E2
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   75
         Picture         =   "BorderoCheque.ctx":466C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   555
         Picture         =   "BorderoCheque.ctx":476E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2025
         Picture         =   "BorderoCheque.ctx":48C8
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1536
         Picture         =   "BorderoCheque.ctx":4A46
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1380
      TabIndex        =   6
      Top             =   255
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   9
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox DataEnvio 
      Height          =   300
      Left            =   4755
      TabIndex        =   8
      Top             =   270
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownDataEnvio 
      Height          =   300
      Left            =   5745
      TabIndex        =   9
      Top             =   270
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.Label LabelTotalNDesc 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   3855
      TabIndex        =   52
      Top             =   5460
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Não Detalhados"
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
      Left            =   3915
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   51
      Top             =   5220
      Width           =   1380
   End
   Begin VB.Label LabelTotalBordero 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   5880
      TabIndex        =   21
      Top             =   5460
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total Borderô"
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
      Left            =   6060
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   20
      Top             =   5220
      Width           =   1170
   End
   Begin VB.Label LabelTotalDesc 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1800
      TabIndex        =   18
      Top             =   5460
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Detalhados"
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
      Left            =   2040
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   17
      Top             =   5220
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data de Envio :"
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
      Index           =   6
      Left            =   3330
      TabIndex        =   10
      Top             =   330
      Width           =   1350
   End
   Begin VB.Label LabelCodigo 
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
      Left            =   630
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   7
      Top             =   315
      Width           =   660
   End
   Begin VB.Label Label3 
      Caption         =   " +                                ="
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
      Left            =   3480
      TabIndex        =   19
      Top             =   5520
      Width           =   2295
   End
End
Attribute VB_Name = "BorderoCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim iAlterado As Integer
Dim objGridCheque As AdmGrid
Dim objGridChequeN As AdmGrid
Dim gcolCheque As Collection
Dim gcolChequeN As Collection

Dim iGrid_Selecionado_Col As Integer
Dim iGrid_Banco_Col As Integer
Dim iGrid_Agencia_Col As Integer
Dim iGrid_Conta_Col As Integer
Dim iGrid_Numero_Col As Integer
Dim iGrid_DataDepositoCheque_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_CPFCGC_Col As Integer

Dim iGrid_SelecionadoN_Col As Integer
Dim iGrid_DataDepositoChequeN_Col As Integer
Dim iGrid_ValorN_Col As Integer

Dim iFrameAtual As Integer

Private WithEvents objEventoBorderoCheque As AdmEvento
Attribute objEventoBorderoCheque.VB_VarHelpID = -1

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objBordero As New ClassBorderoCheque

On Error GoTo Erro_BotaoImprimir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'se o código não estiver preenchido->erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 120060

    'se a data de envio não estiver preenchido->erro
    If Len(Trim(DataEnvio.ClipText)) = 0 Then gError 120061
    
    Call Move_Tela_Memoria(objBordero)
    'If lErro <> SUCESSO Then gError 120062

    lErro = BorderoCheque_Le(objBordero)
    If lErro <> SUCESSO And lErro <> 103966 Then gError 120063
    
    If lErro = 103966 Then gError 120064
    
    '???? adaptar para bordero cheque
    'ver expr. selecao, nome tsk, etc..
    'aguardando tsk ficar pronto....
    'lErro = objRelatorio.ExecutarDireto("Borderô Cheque", "PedidoVenda >= @NPEDVENDINIC E PedidoVenda <= @NPEDVENDFIM", 1, "PedVenda", "NPEDVENDINIC", objPedidoVenda.lCodigo, "NPEDVENDFIM", objPedidoVenda.lCodigo)
    If lErro <> SUCESSO Then gError 120065

    'Limpa a Tela
    Call Limpa_Tela_BorderoCheque

    iAlterado = 0
    
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 120060
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 120061
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
        
        Case 120062, 120063, 120065

        Case 120064
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROCHEQUE_NAOENCONTRADO", gErr, objBordero.iFilialEmpresa, objBordero.lNumBordero)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 143584)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub


End Sub


Public Sub Form_Load()

Dim lErro As Long
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim iIndice As Integer
Dim bCancel As Boolean

On Error GoTo Erro_Form_Load

    iFrameAtual = 1

    'instancia os grids
    Set objGridChequeN = New AdmGrid
    Set objGridCheque = New AdmGrid
    Set objEventoBorderoCheque = New AdmEvento

    'inicializa o grid de cheques detalhados
    lErro = Inicializa_GridCheque(objGridCheque)
    If lErro <> SUCESSO Then gError 103944

    'inicializa o grid de cheques não detalhados
    lErro = Inicializa_GridChequeN(objGridChequeN)
    If lErro <> SUCESSO Then gError 103945

    'carrega a combo de contas bancárias
    lErro = Carrega_Conta_Corrente_Bancaria()
    If lErro <> SUCESSO Then gError 103946

    'preenche o objAdmMeioPagto com os campos para buscar no BD
    objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
    objAdmMeioPagto.iCodigo = MEIO_PAGAMENTO_CHEQUE

    'Lê a admmeiopagto
    lErro = CF("AdmMeioPagto_Le", objAdmMeioPagto)
    If lErro <> SUCESSO And lErro <> 104017 Then gError 103947

    'se não a encontrar-> erro
    If lErro = 104017 Then gError 103948

    'Se contaCorrente interna estiver preenchida
    If objAdmMeioPagto.iContaCorrenteInterna <> 0 Then
 
        'selecioná-la na combo de contas correntes internas
        ContaCorrente.Text = objAdmMeioPagto.iContaCorrenteInterna
        Call ContaCorrente_Validate(bCancel)

    End If

    'preencher a data de envio com a data atual
    DataEnvio.PromptInclude = False
    DataEnvio.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEnvio.PromptInclude = True

    'preencher a data de deposito com a data atual
    DataBomParaAte.PromptInclude = False
    DataBomParaAte.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataBomParaAte.PromptInclude = True

    'se estiver no backofficce
    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then

        'desabilita o nao lhe é pertinente
        BotaoProxNum.Visible = False
        FrameBomPara.Visible = False
        FrameDestino.Enabled = False

    End If
    
    LabelTotalBordero.Caption = Format(0, "STANDARD")
    LabelTotalDesc.Caption = Format(0, "STANDARD")
    LabelTotalNDesc.Caption = Format(0, "STANDARD")

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 103944 To 103947, 103954

        Case 103948
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_CADASTRADO", gErr, objAdmMeioPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143585)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objBorderoCheque As ClassBorderoCheque) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'se objBorderoCheque está definido
    If Not objBorderoCheque Is Nothing Then

        'preenche a tela
        lErro = Traz_BorderoCheque_Tela(objBorderoCheque)
        If lErro <> SUCESSO And lErro <> 103974 Then gError 103978

        'se não encontrou bordero
        If lErro = 103974 Then

            'limpa a tela e preenche o campo do código
            Call Limpa_Tela_BorderoCheque

            Codigo.Text = objBorderoCheque.lNumBordero

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 103978

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143586)

    End Select

    Exit Function

End Function

Private Sub LabelCodigo_Click()

Dim objBorderoCheque As New ClassBorderoCheque
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'se o codigo estiver preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then objBorderoCheque.lNumBordero = StrParaLong(Codigo.Text)

    'chama a tela de browser
    Call Chama_Tela("BorderoChequeLojaLista", colSelecao, objBorderoCheque, objEventoBorderoCheque)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143587)

    End Select

    Exit Sub

End Sub

Private Sub objEventoBorderoCheque_evSelecao(obj1 As Object)

Dim objBorderoCheque As ClassBorderoCheque
Dim lErro As Long
Dim colCheque As New Collection
Dim colChequeN As New Collection

On Error GoTo Erro_objEventoBorderoCheque_evSelecao

    'seta objborderocheque com o recebido por parâmetro
    Set objBorderoCheque = obj1

    'preenche a tela com o borderô selecionado
    lErro = Traz_BorderoCheque_Tela(objBorderoCheque)
    If lErro <> SUCESSO And lErro <> 103974 Then gError 107021

    If lErro = 103974 Then gError 103980

    'mostra a tela preenchida
    Me.Show

    iAlterado = 0

    Exit Sub

Erro_objEventoBorderoCheque_evSelecao:

    Select Case gErr

        Case 103979, 107021

        Case 103980
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROCHEQUE_NAOENCONTRADO", gErr, objBorderoCheque.iFilialEmpresa, objBorderoCheque.lNumBordero)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143588)

    End Select

    Exit Sub

End Sub

Public Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objBorderoCheque As New ClassBorderoCheque

On Error GoTo Erro_Tela_Extrai

    sTabela = "BorderoCheque"

    'preenche o objBorderoCheque com os dados da tela
    Call Move_Tela_Memoria(objBorderoCheque)

    'preenche a coleção de campos-valores
    colCampoValor.Add "DataBackoffice", objBorderoCheque.dtDataBackoffice, 0, "DataBackoffice"
    colCampoValor.Add "DataEnvio", objBorderoCheque.dtDataEnvio, 0, "DataEnvio"
    colCampoValor.Add "DataImpressao", objBorderoCheque.dtDataImpressao, 0, "DataImpressao"
    colCampoValor.Add "CodNossaConta", objBorderoCheque.iCodNossaConta, 0, "CodNossaConta"
    colCampoValor.Add "FilialEmpresa", objBorderoCheque.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "NumBordero", objBorderoCheque.lNumBordero, 0, "NumBordero"

    'estabelece o filtro
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143589)

    End Select

    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objBorderoCheque As New ClassBorderoCheque

On Error GoTo Erro_Tela_Preenche

    'preenche o objBorderoCheque com a colecao de valores
    objBorderoCheque.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objBorderoCheque.lNumBordero = colCampoValor.Item("NumBordero").vValor

    'traz os dados do Deposito Bancario para a tela
    lErro = Traz_BorderoCheque_Tela(objBorderoCheque)
    If lErro <> SUCESSO Then gError 103985

    iAlterado = 0

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 103985

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143590)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'gera um novo código para o borderô
    lErro = BorderoCheque_Codigo_Automatico(lCodigo)
    If lErro <> SUCESSO Then gError 103987

    'preenche o campo com o código gerado
    Codigo.Text = lCodigo

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 103987

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143591)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_GotFocus()

On Error GoTo Erro_Codigo_GotFocus

    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

    Exit Sub

Erro_Codigo_GotFocus:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143592)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'se não estiver preenchido, sai
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub

    'se o código estiver incorreto, sai mantém o foco -> erro
    lErro = Long_Critica(Codigo.Text)
    If lErro <> SUCESSO Then gError 103988

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 103988

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143593)

    End Select

    Exit Sub

End Sub

Private Sub DataEnvio_GotFocus()

On Error GoTo Erro_DataEnvio_GotFocus

    Call MaskEdBox_TrataGotFocus(DataEnvio, iAlterado)

    Exit Sub

Erro_DataEnvio_GotFocus:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143594)

    End Select

    Exit Sub

End Sub

Private Sub DataEnvio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEnvio_Validate

    'se não estiver preenchido-> erro
    If Len(Trim(DataEnvio.ClipText)) = 0 Then Exit Sub

    'critica... se for inválida->erro
    lErro = Data_Critica(DataEnvio.Text)
    If lErro <> SUCESSO Then gError 103989

    Cancel = False

    Exit Sub

Erro_DataEnvio_Validate:

    Cancel = True

    Select Case gErr

        Case 103989

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143595)

    End Select

    Exit Sub

End Sub

Private Sub DataEnvio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownDataEnvio_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvio_DownClick

    'diminui a DataEnvio de um dia
    lErro = Data_Up_Down_Click(DataEnvio, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 103990

    Exit Sub

Erro_UpDownDataEnvio_DownClick:

    Select Case gErr

        Case 103990

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143596)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEnvio_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvio_UpClick

    'diminui a DataEnvio de um dia
    lErro = Data_Up_Down_Click(DataEnvio, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 103991

    Exit Sub

Erro_UpDownDataEnvio_UpClick:

    Select Case gErr

        Case 103991

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143597)

    End Select

    Exit Sub

End Sub

Private Sub ContaCorrente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaCorrente_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaCorrente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_ContaCorrente_Validate

    'se estiver preenchida
    If Len(Trim(ContaCorrente.Text)) > 0 Then

        'seleciona o elemento da combo
        lErro = Combo_Seleciona(ContaCorrente, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 103992

        'se não existir pelo código ->erro
        If lErro = 6730 Then gError 103993

        'se não existir pela string-> erro
        If lErro = 6731 Then gError 103994

    End If

    Exit Sub

Erro_ContaCorrente_Validate:

    Cancel = True

    Select Case gErr

        Case 103992

        Case 103993, 103994
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTASCORRENTESINTERNAS_NAOENCONTRADA", gErr, ContaCorrente.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143598)

    End Select

    Exit Sub

End Sub

Private Sub DataBomParaDe_GotFocus()

On Error GoTo Erro_DataBomParaDe_GotFocus

    Call MaskEdBox_TrataGotFocus(DataBomParaDe, iAlterado)

    Exit Sub

Erro_DataBomParaDe_GotFocus:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143599)

    End Select

    Exit Sub

End Sub

Private Sub DataBomParaDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataBomParaDe_Validate

    'se estiver preenhida
    If Len(Trim(DataBomParaDe.ClipText)) > 0 Then

        'critica. se for inválida->erro
        lErro = Data_Critica(DataBomParaDe.Text)
        If lErro <> SUCESSO Then gError 103995

    End If

    Exit Sub

Erro_DataBomParaDe_Validate:

    Cancel = True

    Select Case gErr

        Case 103995

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143600)

    End Select

    Exit Sub

End Sub

Private Sub DataBomParaDe_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDe_DownClick

    'diminui a DataEnvio de um dia
    lErro = Data_Up_Down_Click(DataBomParaDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 103996

    Exit Sub

Erro_UpDownDe_DownClick:

    Select Case gErr

        Case 103996

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143601)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDe_UpClick

    'diminui a DataEnvio de um dia
    lErro = Data_Up_Down_Click(DataBomParaDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 103997

    Exit Sub

Erro_UpDownDe_UpClick:

    Select Case gErr

        Case 103997

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143602)

    End Select

    Exit Sub

End Sub

Private Sub DataBomParaAte_GotFocus()

On Error GoTo Erro_DataBomParaAte_GotFocus

    Call MaskEdBox_TrataGotFocus(DataBomParaAte, iAlterado)

    Exit Sub

Erro_DataBomParaAte_GotFocus:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143603)

    End Select

    Exit Sub

End Sub

Private Sub DataBomParaAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataBomParaAte_Validate

    'se estiver preenchida
    If Len(Trim(DataBomParaAte.ClipText)) > 0 Then

        'critica. se inválida->erro
        lErro = Data_Critica(DataBomParaAte.Text)
        If lErro <> SUCESSO Then gError 103998

    End If

    Exit Sub

Erro_DataBomParaAte_Validate:

    Cancel = True

    Select Case gErr

        Case 103998

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143604)

    End Select

    Exit Sub

End Sub

Private Sub DataBomParaAte_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownAte_DownClick

    'diminui a DataBomParaAte de um dia
    lErro = Data_Up_Down_Click(DataBomParaAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 107000

    Exit Sub

Erro_UpDownAte_DownClick:

    Select Case gErr

        Case 107000

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143605)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownAte_UpClick

    'diminui a DataBomParaAte de um dia
    lErro = Data_Up_Down_Click(DataBomParaAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 107001

    Exit Sub

Erro_UpDownAte_UpClick:

    Select Case gErr

        Case 107001

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143606)

    End Select

    Exit Sub

End Sub

Private Sub Valor_GotFocus()

On Error GoTo Erro_Valor_GotFocus

    Call MaskEdBox_TrataGotFocus(Valor, iAlterado)

    Exit Sub

Erro_Valor_GotFocus:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143607)

    End Select

    Exit Sub

End Sub

Private Sub OptionConta_Click()

    'habilita a label de contacorrente e a combo referente
    LabelContaCorrente.Enabled = True
    ContaCorrente.Enabled = True

End Sub

Private Sub OptionBackoffice_Click()

    'limpa a combo de contacorrente, desabilita a combo e a label
    ContaCorrente.ListIndex = -1
    ContaCorrente.Enabled = False
    LabelContaCorrente.Enabled = False

End Sub

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        FrameCheque(TabStrip1.SelectedItem.index).Visible = True

        'Torna Frame atual visivel
        FrameCheque(iFrameAtual).Visible = False

        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.index

    End If

End Sub

Private Sub GridCheques_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCheque, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCheque, iAlterado)
    End If

End Sub

Private Sub GridCheques_EnterCell()

    Call Grid_Entrada_Celula(objGridCheque, iAlterado)

End Sub

Private Sub GridCheques_GotFocus()

    Call Grid_Recebe_Foco(objGridCheque)

End Sub

Private Sub GridCheques_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridCheque)

End Sub

Private Sub GridCheques_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCheque, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCheque, iAlterado)
    End If

End Sub

Private Sub GridCheques_LeaveCell()

    Call Saida_Celula(objGridCheque)

End Sub

Private Sub GridCheques_LostFocus()

    Call Grid_Libera_Foco(objGridCheque)

End Sub

Private Sub GridCheques_RowColChange()

    Call Grid_RowColChange(objGridCheque)

End Sub

Private Sub GridCheques_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridCheque)

End Sub

Private Sub Selecionado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCheque)

End Sub

Private Sub Selecionado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCheque)

End Sub

Private Sub Selecionado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCheque.objControle = Selecionado
    lErro = Grid_Campo_Libera_Foco(objGridCheque)

    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Selecionado_Click()

On Error GoTo Erro_Selecionado_Click

    'se a linha atual estiver desmarcada
    If StrParaInt(GridCheques.TextMatrix(GridCheques.Row, iGrid_Selecionado_Col)) = DESMARCADO Then

        'soma o cheque selecionado aos totalizadores respectivos
        LabelTotalDesc.Caption = Format((StrParaDbl(LabelTotalDesc.Caption) - StrParaDbl(GridCheques.TextMatrix(GridCheques.Row, iGrid_Valor_Col))), "STANDARD")
        LabelTotalBordero.Caption = Format((StrParaDbl(LabelTotalBordero.Caption) - StrParaDbl(GridCheques.TextMatrix(GridCheques.Row, iGrid_Valor_Col))), "STANDARD")

    'se a linha atual estiver marcada
    Else

        'subtrai o cheque selecionado aos totalizadores respectivos
        LabelTotalDesc.Caption = Format((StrParaDbl(LabelTotalDesc.Caption) + StrParaDbl(GridCheques.TextMatrix(GridCheques.Row, iGrid_Valor_Col))), "STANDARD")
        LabelTotalBordero.Caption = Format((StrParaDbl(LabelTotalBordero.Caption) + StrParaDbl(GridCheques.TextMatrix(GridCheques.Row, iGrid_Valor_Col))), "STANDARD")

    End If

    Exit Sub

Erro_Selecionado_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143608)

    End Select

    Exit Sub

End Sub

Private Sub GridChequeN_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridChequeN, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridChequeN, iAlterado)
    End If

End Sub

Private Sub GridChequeN_EnterCell()

    Call Grid_Entrada_Celula(objGridChequeN, iAlterado)

End Sub

Private Sub GridChequeN_GotFocus()

    Call Grid_Recebe_Foco(objGridChequeN)

End Sub

Private Sub GridChequeN_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridChequeN)

End Sub

Private Sub GridChequeN_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridChequeN, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridChequeN, iAlterado)
    End If

End Sub

Private Sub GridChequeN_LeaveCell()

    Call Saida_Celula(objGridChequeN)

End Sub

Private Sub GridChequeN_LostFocus()

    Call Grid_Libera_Foco(objGridChequeN)

End Sub

Private Sub GridChequeN_RowColChange()

    Call Grid_RowColChange(objGridChequeN)

End Sub

Private Sub GridChequeN_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridChequeN)

End Sub

Private Sub SelecionadoN_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridChequeN)

End Sub

Private Sub SelecionadoN_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridChequeN)

End Sub

Private Sub SelecionadoN_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridChequeN.objControle = SelecionadoN
    lErro = Grid_Campo_Libera_Foco(objGridChequeN)

    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub SelecionadoN_Click()

On Error GoTo Erro_SelecionadoN_Click

    'se a linha atual estiver desmarcada
    If StrParaInt(GridChequeN.TextMatrix(GridChequeN.Row, iGrid_SelecionadoN_Col)) = DESMARCADO Then

        'soma o cheque selecionado aos totalizadores respectivos
        LabelTotalNDesc.Caption = Format((StrParaDbl(LabelTotalNDesc.Caption) - StrParaDbl(GridChequeN.TextMatrix(GridChequeN.Row, iGrid_ValorN_Col))), "STANDARD")
        LabelTotalBordero.Caption = Format((StrParaDbl(LabelTotalBordero.Caption) - StrParaDbl(GridChequeN.TextMatrix(GridChequeN.Row, iGrid_ValorN_Col))), "STANDARD")

    'se a linha atual estiver marcada
    Else

        'subtrai o cheque selecionado aos totalizadores respectivos
        LabelTotalNDesc.Caption = Format((StrParaDbl(LabelTotalNDesc.Caption) + StrParaDbl(GridChequeN.TextMatrix(GridChequeN.Row, iGrid_ValorN_Col))), "STANDARD")
        LabelTotalBordero.Caption = Format((StrParaDbl(LabelTotalBordero.Caption) + StrParaDbl(GridChequeN.TextMatrix(GridChequeN.Row, iGrid_ValorN_Col))), "STANDARD")

    End If

    Exit Sub

Erro_SelecionadoN_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143609)

    End Select

    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 107002

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 107002
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143610)

    End Select

    Exit Function

End Function

Private Sub BotaoTrazer_Click()

Dim lErro As Long
Dim dtDataInicio As Date
Dim dtDataFim As Date
Dim colCheque As New Collection
Dim colChequeN As New Collection
Dim iIndice As Integer
Dim vbMsgResp As VbMsgBoxResult

On Error GoTo Erro_BotaoTrazer_Click

    'se a data de início estiver preenchida
    If Len(Trim(DataBomParaDe.ClipText)) <> 0 Then dtDataInicio = StrParaDate(DataBomParaDe.Text)

    'se a data de dim estiver preenchida
    If Len(Trim(DataBomParaAte.ClipText)) <> 0 Then dtDataFim = StrParaDate(DataBomParaAte.Text)

    'se ambas as datas estiverem preenchidas
    If Len(Trim(DataBomParaDe.ClipText)) <> 0 And Len(Trim(DataBomParaAte.ClipText)) <> 0 Then

        'se a data de início for maior que a de fim->erro
        If dtDataInicio > dtDataFim Then gError 107017

    End If

    'lê os cheques que atendem ao intervalo preenchido
    lErro = CF("Cheque_Le1", colCheque, colChequeN, dtDataInicio, dtDataFim)
    If lErro <> SUCESSO And lErro <> 107010 Then gError 107013

    'se a seleção retornou vazia -> erro
    If lErro = 107010 Then gError 107014

    'se algum grid já estiver preenchido
    If objGridCheque.iLinhasExistentes <> 0 Or objGridChequeN.iLinhasExistentes <> 0 Then

        'pergunta se deseja apagar o grid ou adicionar ao fim do que já está preenchido
        vbMsgResp = Rotina_Aviso(vbYesNo, "AVISO_LIMPAR_GRID")

        'se a resposta for sim
        If vbMsgResp = vbYes Then

            'apaga o grid de cheques especificados
            lErro = Grid_Limpa(objGridCheque)
            If lErro <> SUCESSO Then gError 107018

            'apaga o grid de cheques não especificados
            lErro = Grid_Limpa(objGridChequeN)
            If lErro <> SUCESSO Then gError 107019

            Set gcolCheque = Nothing
            Set gcolChequeN = Nothing

            'limpa os totais
            LabelTotalDesc.Caption = Format(0, "STANDARD")
            LabelTotalNDesc.Caption = Format(0, "STANDARD")
            LabelTotalBordero.Caption = Format(0, "STANDARD")

        End If

    End If

    'preenche o grid de cheques especificados com os cheques especificados
    lErro = GridCheque_Preenche(colCheque)
    If lErro <> SUCESSO Then gError 107015

    'preenche o grid de cheques não especificados com os cheques não especificados
    lErro = GridChequeN_Preenche(colChequeN)
    If lErro <> SUCESSO Then gError 107016

    Exit Sub

Erro_BotaoTrazer_Click:

    Select Case gErr

        Case 107017
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAFIM_MAIOR_DATAINICIO", gErr)

        Case 107013, 103015, 107016 To 107019

        Case 107014
            If dtDataInicio <> 0 And dtDataFim <> 0 Then
                Call Rotina_Aviso(vbOKOnly, "AVISO_CHEQUEPRE_VAZIA_2", dtDataInicio, dtDataFim)
            Else
                If dtDataInicio <> 0 Then
                    Call Rotina_Aviso(vbOKOnly, "AVISO_CHEQUEPRE_VAZIA_1A", dtDataInicio)
                Else
                    If dtDataFim <> 0 Then
                        Call Rotina_Aviso(vbOKOnly, "AVISO_CHEQUEPRE_VAZIA_1B", dtDataFim)
                    Else
                        Call Rotina_Aviso(vbOKOnly, "AVISO_CHEQUEPRE_VAZIA_0")
                    End If
                End If
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143611)

    End Select

    Exit Sub

End Sub

Private Sub BotaoMarcarTodos_Click()

Dim iIndice As Integer
Dim lErro As Long
Dim dTotal As Double

On Error GoTo Erro_BotaoMarcarTodos_Click

    'varre o grid de cheques
    For iIndice = 1 To objGridCheque.iLinhasExistentes

        'marca cada um e acumulando a soma
        GridCheques.TextMatrix(iIndice, iGrid_Selecionado_Col) = MARCADO
        dTotal = dTotal + StrParaDbl(GridCheques.TextMatrix(iIndice, iGrid_Valor_Col))

    Next

    'dá um refresh nas checkboxes
    lErro = Grid_Refresh_Checkbox(objGridCheque)
    If lErro <> SUCESSO Then gError 107023

    'atualiza os totalizadores
    LabelTotalDesc.Caption = Format(dTotal, "STANDARD")
    LabelTotalBordero.Caption = Format((StrParaDbl(LabelTotalDesc.Caption) + StrParaDbl(LabelTotalNDesc.Caption)), "STANDARD")

    Exit Sub

Erro_BotaoMarcarTodos_Click:

    Select Case gErr

        Case 107023

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143612)

    End Select

    Exit Sub

End Sub

Private Sub BotaoDesmarcarTodos_Click()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoDesmarcarTodos_Click

    'varre o grid de cheques
    For iIndice = 1 To objGridCheque.iLinhasExistentes

        'desmarcando cada uma
        GridCheques.TextMatrix(iIndice, iGrid_Selecionado_Col) = DESMARCADO

    Next

    'dá um refresh nas checkboxes
    lErro = Grid_Refresh_Checkbox(objGridCheque)
    If lErro <> SUCESSO Then gError 107024

    'atualiza os totalizadores
    LabelTotalDesc.Caption = Format(0, "STANDARD")
    LabelTotalBordero.Caption = LabelTotalNDesc.Caption

    Exit Sub

Erro_BotaoDesmarcarTodos_Click:

    Select Case gErr

        Case 107024

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143613)

    End Select

    Exit Sub

End Sub

Private Sub BotaoMarcarTodosN_Click()

Dim iIndice As Integer
Dim lErro As Long
Dim dTotal As Double

On Error GoTo Erro_BotaoMarcarTodosN_Click

    'varre o grid de cheques
    For iIndice = 1 To objGridChequeN.iLinhasExistentes

        'marca cada um e acumulando a soma
        GridChequeN.TextMatrix(iIndice, iGrid_SelecionadoN_Col) = MARCADO
        dTotal = dTotal + StrParaDbl(GridChequeN.TextMatrix(iIndice, iGrid_ValorN_Col))

    Next

    'dá um refresh nas checkboxes
    lErro = Grid_Refresh_Checkbox(objGridChequeN)
    If lErro <> SUCESSO Then gError 107023

    'atualiza os totalizadores
    LabelTotalNDesc.Caption = Format(dTotal, "STANDARD")
    LabelTotalBordero.Caption = Format((StrParaDbl(LabelTotalDesc.Caption) + StrParaDbl(LabelTotalNDesc.Caption)), "STANDARD")

    Exit Sub

Erro_BotaoMarcarTodosN_Click:

    Select Case gErr

        Case 107023

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143614)

    End Select

    Exit Sub

End Sub

Private Sub BotaoDesmarcarTodosN_Click()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoDesmarcarTodosN_Click

    'varre o grid de cheques
    For iIndice = 1 To objGridChequeN.iLinhasExistentes

        'desmarca cada um
        GridChequeN.TextMatrix(iIndice, iGrid_SelecionadoN_Col) = DESMARCADO

    Next

    'dá um refresh nas checkboxes
    lErro = Grid_Refresh_Checkbox(objGridChequeN)
    If lErro <> SUCESSO Then gError 107024

    'atualiza os totalizadores
    LabelTotalNDesc.Caption = Format(0, "STANDARD")
    LabelTotalBordero.Caption = LabelTotalDesc.Caption

    Exit Sub

Erro_BotaoDesmarcarTodosN_Click:

    Select Case gErr

        Case 107024

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143615)

    End Select

    Exit Sub

End Sub

Private Sub BotaoEditarN_Click()

Dim objChequePre As New ClassChequePre

On Error GoTo Erro_BotaoEditarN_Click

    'verifica se tem alguma llinha selecionada
    If GridChequeN.Row <= 0 Then gError 107026

    'seta um chequepre com o elemento da coleção global correspondente ao do grid
    Set objChequePre = gcolChequeN.Item(GridChequeN.Row)

    'chama a tela de cheques
    Call Chama_Tela("ChequeNEsp", objChequePre)

    Exit Sub

Erro_BotaoEditarN_Click:

    Select Case gErr

        Case 107026
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143616)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a gravar registro
    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 107085

    Call Limpa_Tela_BorderoCheque

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 107085

            Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143617)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim objBorderoCheque As New ClassBorderoCheque
Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'se o código não estiver preenchido->erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 107079

    'se a data de envio não estiver preenchido->erro
    If Len(Trim(DataEnvio.ClipText)) = 0 Then gError 107080

    'se a contacorrente estiver marcada
    If OptionConta.Value = True Then
        'se não houver conta selecionada->erro
        If ContaCorrente.ListIndex = -1 Then gError 107081

    End If

    'se o total estiver zerado significa que não há cheque selecionado->erro
    If StrParaDbl(LabelTotalBordero.Caption) = 0 Then gError 107082

    'move da tela para memória
    Call Move_Tela_Memoria(objBorderoCheque)

    'preenche as coleções de cheques com os cheques selecionados nos 2 grids
    Call Move_Tela_Memoria_Cheque(objBorderoCheque.colCheque, objBorderoCheque.colChequeN)

    'preenche a coleção de cheques não
    lErro = Trata_Alteracao(objBorderoCheque, objBorderoCheque.iFilialEmpresa, objBorderoCheque.lNumBordero)
    If lErro <> SUCESSO Then gError 107084

    'chama a função de gravação de borderô
    lErro = CF("BorderoCheque_Grava", objBorderoCheque)
    If lErro <> SUCESSO Then gError 107083
    
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 107079
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 107080
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case 107081
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", gErr)

        Case 107082
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROCHEQUE_ZERO", gErr)

        Case 107083, 107084

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143618)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgResp As VbMsgBoxResult
Dim objBorderoCheque As New ClassBorderoCheque

On Error GoTo Erro_BotaoExcluir_Click

    'se o código não estiver preenchido-> erro
    If Len(Trim(Codigo)) = 0 Then gError 107104

    'preenche os atributos para buscar um determinado bordero de cheque
    objBorderoCheque.iFilialEmpresa = giFilialEmpresa
    objBorderoCheque.lNumBordero = StrParaLong(Codigo.Text)

    'busca o bordero
    lErro = BorderoCheque_Le(objBorderoCheque)
    If lErro <> SUCESSO And lErro <> 103966 Then gError 107105

    'se ele não existir-> erro
    If lErro = 103966 Then gError 107106

    'pergunta se tem certeza que deseja excluir
    vbMsgResp = Rotina_Aviso(vbYesNo, "AVISO_BORDEROCHEQUE_EXCLUSAO", objBorderoCheque.iFilialEmpresa, objBorderoCheque.lNumBordero)

    'se sim
    If vbMsgResp = vbYes Then

        'exclui o bordero
        lErro = CF("BorderoCheque_Exclui", objBorderoCheque)
        If lErro <> SUCESSO Then gError 107107

        'limpa a tela
        Call Limpa_Tela_BorderoCheque

        iAlterado = 0

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 107104
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 107105, 107107

        Case 107106
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROCHEQUE_NAOENCONTRADO", gErr, objBorderoCheque.iFilialEmpresa, objBorderoCheque.lNumBordero)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143619)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'testa se há alguma alteração na tela
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 107108

    'limpra a tela
    Call Limpa_Tela_BorderoCheque

    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 107109

    iAlterado = 0

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 107108, 107109

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143620)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_BorderoCheque()

Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_BorderoCheque

    'limpa a tela
    Call Limpa_Tela(Me)

    'limpa o grid do primeiro tab
    lErro = Grid_Limpa(objGridCheque)
    If lErro <> SUCESSO Then gError 103949

    'limpa o grid do segundo tab
    lErro = Grid_Limpa(objGridChequeN)
    If lErro <> SUCESSO Then gError 103950

    'limpa as coleções globais
    Set gcolChequeN = Nothing
    Set gcolCheque = Nothing

    'preenche os campos da administradora para busca no bd
    objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
    objAdmMeioPagto.iCodigo = MEIO_PAGAMENTO_CHEQUE

    'lê a administradora de meio de pagamento
    lErro = CF("AdmMeioPagto_Le", objAdmMeioPagto)
    If lErro <> SUCESSO And lErro <> 104017 Then gError 103951

    'se não encontrou-> erro
    If lErro = 104017 Then gError 103952

    'se o admmeiopagto veio com conta interna preenchida
    If objAdmMeioPagto.iContaCorrenteInterna <> 0 Then

        'escolhe a da combo
        Call Combo_Seleciona_ItemData(ContaCorrente, objAdmMeioPagto.iContaCorrenteInterna)

    End If

    'preenche as datas de envio e bomparaate com a data atual
    DataEnvio.PromptInclude = False
    DataEnvio.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEnvio.PromptInclude = True

    DataBomParaAte.PromptInclude = False
    DataBomParaAte.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataBomParaAte.PromptInclude = True

    'zera os totais
    LabelTotalBordero.Caption = Format(0, "STANDARD")
    LabelTotalDesc.Caption = Format(0, "STANDARD")
    LabelTotalNDesc.Caption = Format(0, "STANDARD")

    Exit Sub

Erro_Limpa_Tela_BorderoCheque:

    Select Case gErr

        Case 103949 To 103951, 103953

        Case 103952
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_CADASTRADO", gErr, objAdmMeioPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143621)

    End Select

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub form_unload(Cancel As Integer)

On Error GoTo Erro_Form_Unload

    'libera as variáveis globais
    Set objGridCheque = Nothing
    Set objGridChequeN = Nothing
    Set gcolCheque = Nothing
    Set gcolChequeN = Nothing

    'libera o comando de setas
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143622)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Function Traz_BorderoCheque_Tela(objBorderoCheque As ClassBorderoCheque) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_BorderoCheque_Tela

    'Limpa a tela
    Call Limpa_Tela_BorderoCheque

    'Tenta le o bordero selecionado
    lErro = BorderoCheque_Le(objBorderoCheque)
    If lErro <> SUCESSO And lErro <> 103966 And lErro <> 103968 Then gError 103973

    'se o bordero não foi encontrado->erro
    If lErro = 103966 Then gError 103974

    'se não foram encontrados cheques vinculados ao bordero-> erro
    If lErro = 103968 Then gError 107057

    'Preenche o código
    Codigo.Text = objBorderoCheque.lNumBordero

    'preenche a data de envio
    If objBorderoCheque.dtDataEnvio <> DATA_NULA Then

        DataEnvio.PromptInclude = False
        DataEnvio.Text = Format(objBorderoCheque.dtDataEnvio, "dd/mm/yy")
        DataEnvio.PromptInclude = True

    End If

    'marca uma das opções entre conta corrente e back office
    If objBorderoCheque.iCodNossaConta = 0 Then

        'se for backoffice
        OptionBackoffice.Value = True

    Else

        'se não for backoffice
        OptionConta.Value = True

        'seleciona a conta corrente na combo
        Call Combo_Seleciona_ItemData(ContaCorrente, objBorderoCheque.iCodNossaConta)

    End If

    'preenche o grid de cheques detalhados
    lErro = GridCheque_Preenche(objBorderoCheque.colCheque)
    If lErro <> SUCESSO Then gError 103976

    'preenche o grid de cheques não detalhados
    lErro = GridChequeN_Preenche(objBorderoCheque.colChequeN)
    If lErro <> SUCESSO Then gError 103977

    Traz_BorderoCheque_Tela = SUCESSO

    Exit Function

Erro_Traz_BorderoCheque_Tela:

    Traz_BorderoCheque_Tela = gErr

    Select Case gErr

        Case 103973 To 103977, 107057

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143623)

    End Select

    Exit Function

End Function

Public Function BorderoCheque_Le(ByVal objBorderoCheque As ClassBorderoCheque) As Long
'Função que recebe um objBorderoCheque com os campos filialempresa e codigo preenchidos
'e o retorna preenchido com a coleção de cheques, inclusive, caso encontre

Dim lErro As Long
Dim colCheque As New Collection
Dim colChequeN As New Collection

On Error GoTo Erro_BorderoCheque_Le

    'le o bordero sem preencher sua coleção
    lErro = CF("BorderoCheque_Le1", objBorderoCheque)
    If lErro <> SUCESSO And lErro <> 103959 Then gError 103965

    'se não encontrar, erro
    If lErro = 103959 Then gError 103966

    'Lê os cheques referentes ao borderô encontrado
    lErro = CF("Cheque_Le_Bordero", objBorderoCheque)
    If lErro <> SUCESSO And lErro <> 103963 Then gError 103967

    'se não encontrar nenhum cheque, erro
    If lErro = 103963 Then gError 103968

    BorderoCheque_Le = SUCESSO

    Exit Function

Erro_BorderoCheque_Le:

    BorderoCheque_Le = gErr

    Select Case gErr

        Case 103965 To 103968
            'serão tratados na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143624)

    End Select

    Exit Function

End Function

Private Function Carrega_Conta_Corrente_Bancaria() As Long

Dim colCodigoNomeConta As New AdmColCodigoNome
Dim objCodigoNomeConta As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_Conta_Corrente_Bancaria

    'Carrega a coleção de contas
    lErro = CF("ContasCorrentes_Bancarias_Le_CodigosNomesRed", colCodigoNomeConta)
    If lErro <> SUCESSO Then gError 103971

    'se retornar a coleção vazia->erro
    If colCodigoNomeConta.Count = 0 Then gError 103972

    'adiciona cada conta corrente da coleção à combo
    For Each objCodigoNomeConta In colCodigoNomeConta

        ContaCorrente.AddItem CStr(objCodigoNomeConta.iCodigo) & SEPARADOR & objCodigoNomeConta.sNome
        ContaCorrente.ItemData(ContaCorrente.NewIndex) = objCodigoNomeConta.iCodigo

    Next

    Carrega_Conta_Corrente_Bancaria = SUCESSO

    Exit Function

Erro_Carrega_Conta_Corrente_Bancaria:

    Carrega_Conta_Corrente_Bancaria = gErr

    Select Case gErr

        Case 103971

        Case 103972
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTASCORRENTESINTERNAS_VAZIA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143625)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridCheque(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_GridCheque

    'tela em Questão
    Set objGridInt.objForm = Me

    With objGridInt

        'Títulos do Grid
        .colColuna.Add ""
        .colColuna.Add "Sel."
        .colColuna.Add "Banco"
        .colColuna.Add "Agência"
        .colColuna.Add "Conta"
        .colColuna.Add "Número"
        .colColuna.Add "Bom Para"
        .colColuna.Add "Valor"
        .colColuna.Add "Cliente"

        'campos vinculados ao Grid
        .colCampo.Add Selecionado.Name
        .colCampo.Add Banco.Name
        .colCampo.Add Agencia.Name
        .colCampo.Add Conta.Name
        .colCampo.Add Numero.Name
        .colCampo.Add DataDepositoCheque.Name
        .colCampo.Add Valor.Name
        .colCampo.Add CPFCGC.Name

        'numeração das colunas do grid
        iGrid_Selecionado_Col = 1
        iGrid_Banco_Col = 2
        iGrid_Agencia_Col = 3
        iGrid_Conta_Col = 4
        iGrid_Numero_Col = 5
        iGrid_DataDepositoCheque_Col = 6
        iGrid_Valor_Col = 7
        iGrid_CPFCGC_Col = 8

        .objGrid = GridCheques
        .iGridLargAuto = GRID_LARGURA_AUTOMATICA
        .iLinhasVisiveis = 5

        .iProibidoIncluir = GRID_PROIBIDO_INCLUIR
        .iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    End With

    GridCheques.ColWidth(0) = 400
    GridCheques.Rows = 6

    Call Grid_Inicializa(objGridInt)

    Inicializa_GridCheque = SUCESSO

    Exit Function

Erro_Inicializa_GridCheque:

    Inicializa_GridCheque = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143626)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridChequeN(objGridInt As AdmGrid) As Long

On Error GoTo Erro_Inicializa_GridChequeN

    'tela em Questão
    Set objGridInt.objForm = Me

    With objGridInt

        'Colunas do Grid
        .colColuna.Add ""
        .colColuna.Add "Selecionado"
        .colColuna.Add "Bom Para"
        .colColuna.Add "Valor"

        'Campos do Grid
        .colCampo.Add SelecionadoN.Name
        .colCampo.Add DataDepositoChequeN.Name
        .colCampo.Add ValorN.Name

        .objGrid = GridChequeN
        .iGridLargAuto = GRID_LARGURA_AUTOMATICA
        .iLinhasVisiveis = 5

        .iProibidoIncluir = GRID_PROIBIDO_INCLUIR
        .iProibidoExcluir = GRID_PROIBIDO_EXCLUIR


    End With

    'numeração das colunas
    iGrid_SelecionadoN_Col = 1
    iGrid_DataDepositoChequeN_Col = 2
    iGrid_ValorN_Col = 3

    GridChequeN.ColWidth(0) = 400
    GridChequeN.Rows = 6

    Call Grid_Inicializa(objGridInt)

    Inicializa_GridChequeN = SUCESSO

    Exit Function

Erro_Inicializa_GridChequeN:

    Inicializa_GridChequeN = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143627)

    End Select

    Exit Function

End Function

Private Function GridCheque_Preenche(colCheque As Collection) As Long

Dim objChequePre As ClassChequePre
Dim objChequePreExistente As ClassChequePre
Dim iLinha As Integer
Dim dTotal As Double
Dim iIndice As Integer
Dim iLinhaInicial As Integer
Dim sTextoFormatado As String
Dim lErro As Long

On Error GoTo Erro_GridCheque_Preenche

    'inicializa a partir que qual linha o grid deve ser preenchido
    iLinhaInicial = objGridCheque.iLinhasExistentes + 1

    'inicializa o total com o total atual
    dTotal = StrParaDbl(LabelTotalDesc.Caption)

    'se ainda não existe uma coleção global de cheques, cria uma
    If gcolCheque Is Nothing Then Set gcolCheque = New Collection

    'adiciona cada cheque da coleção local à coleção global
    For Each objChequePre In colCheque

        iIndice = 0

        'busca o cheque da coleção local na coleção global
        For Each objChequePreExistente In gcolCheque

            If objChequePreExistente.iFilialEmpresaLoja = objChequePre.iFilialEmpresaLoja And _
            objChequePreExistente.lSequencialLoja = objChequePre.lSequencialLoja Then Exit For

            iIndice = iIndice + 1

        Next

        'se não encontrou, adiciona
        If iIndice = gcolCheque.Count Then gcolCheque.Add objChequePre

    Next

    'se existirem mais cheques na coleção do que linhas disponíveis, acerta a quantidade de linhas
    If gcolCheque.Count > GridCheques.Rows - 1 Then

        GridCheques.Rows = gcolCheque.Count + 1

        'reinicializa o grid
        Call Grid_Inicializa(objGridCheque)

    End If

    'preenche as linhas do grid
    For iIndice = iLinhaInicial To gcolCheque.Count

        Set objChequePre = gcolCheque.Item(iIndice)

        GridCheques.TextMatrix(iIndice, iGrid_Banco_Col) = objChequePre.iBanco
        GridCheques.TextMatrix(iIndice, iGrid_Agencia_Col) = objChequePre.sAgencia
        GridCheques.TextMatrix(iIndice, iGrid_Conta_Col) = objChequePre.sContaCorrente
        GridCheques.TextMatrix(iIndice, iGrid_Numero_Col) = objChequePre.lNumero
        GridCheques.TextMatrix(iIndice, iGrid_DataDepositoCheque_Col) = Format(objChequePre.dtDataDeposito, "dd/mm/yyyy")
        GridCheques.TextMatrix(iIndice, iGrid_Valor_Col) = Format(objChequePre.dValor, "STANDARD")

        lErro = Formata_CNPJ_CPF(sTextoFormatado, objChequePre.sCPFCGC)
        If lErro <> SUCESSO Then gError 107058

        GridCheques.TextMatrix(iIndice, iGrid_CPFCGC_Col) = sTextoFormatado

        'acumula o cheque atual no total
        dTotal = dTotal + objChequePre.dValor

    Next

    'acerta a quantidade de linhas existentes no grid
    objGridCheque.iLinhasExistentes = gcolCheque.Count

    'atualiza as checkboxes do grid
    lErro = Grid_Refresh_Checkbox(objGridCheque)
    If lErro <> SUCESSO Then gError 103969

    'Atualiza as labels totalizadoras
    LabelTotalDesc = Format(dTotal, "STANDARD")
    LabelTotalBordero = Format((StrParaDbl(LabelTotalNDesc.Caption) + StrParaDbl(LabelTotalDesc.Caption)), "STANDARD")
    
    Call BotaoMarcarTodos_Click

    GridCheque_Preenche = SUCESSO

    Exit Function

Erro_GridCheque_Preenche:

    GridCheque_Preenche = gErr

    Select Case gErr

        Case 103969, 107058

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143628)

    End Select

    Exit Function

End Function

Private Function GridChequeN_Preenche(colChequeN As Collection) As Long

Dim objChequePre As ClassChequePre
Dim objChequePreExistente As ClassChequePre
Dim iLinha As Integer
Dim dTotal As Double
Dim iIndice As Integer
Dim iLinhaInicial As Integer
Dim sTextoFormatado As String
Dim lErro As Long

On Error GoTo Erro_GridChequeN_Preenche

    'inicializa a partir que qual linha o grid deve ser preenchido
    iLinhaInicial = objGridChequeN.iLinhasExistentes + 1

    'inicializa o total com o total atual
    dTotal = StrParaDbl(LabelTotalNDesc.Caption)

    'se ainda não existe uma coleção global de cheques, cria uma
    If gcolChequeN Is Nothing Then Set gcolChequeN = New Collection

    'adiciona cada cheque da coleção local à coleção global
    For Each objChequePre In colChequeN

        iIndice = 0

        'busca o cheque da coleção local na coleção global
        For Each objChequePreExistente In gcolChequeN

            If objChequePreExistente.iFilialEmpresaLoja = objChequePre.iFilialEmpresaLoja And _
            objChequePreExistente.lSequencialLoja = objChequePre.lSequencialLoja Then Exit For

            iIndice = iIndice + 1

        Next

        'se não encontrou, adiciona
        If iIndice = gcolChequeN.Count Then gcolChequeN.Add objChequePre

    Next

    'se existirem mais cheques na coleção do que linhas disponíveis, acerta a quantidade de linhas
    If gcolChequeN.Count > GridChequeN.Rows - 1 Then

            GridChequeN.Rows = gcolChequeN.Count + 1

            'reinicializa o grid
            Call Grid_Inicializa(objGridChequeN)

    End If

    'preenche as linhas do grid
    For iIndice = iLinhaInicial To gcolChequeN.Count

        Set objChequePre = gcolChequeN.Item(iIndice)

        GridChequeN.TextMatrix(iIndice, iGrid_DataDepositoChequeN_Col) = Format(objChequePre.dtDataDeposito, "dd/mm/yyyy")
        GridChequeN.TextMatrix(iIndice, iGrid_ValorN_Col) = Format(objChequePre.dValor, "STANDARD")

        'acumula o cheque atual no total
        dTotal = dTotal + objChequePre.dValor

    Next

    'acerta a quantidade de linhas existentes no grid
    objGridChequeN.iLinhasExistentes = gcolChequeN.Count

    'atualiza as checkboxes do grid
    lErro = Grid_Refresh_Checkbox(objGridChequeN)
    If lErro <> SUCESSO Then gError 107059

    'Atualiza as labels totalizadoras
    LabelTotalNDesc = Format(dTotal, "STANDARD")
    LabelTotalBordero = Format((StrParaDbl(LabelTotalNDesc.Caption) + StrParaDbl(LabelTotalDesc.Caption)), "STANDARD")
    
    Call BotaoMarcarTodosN_Click

    GridChequeN_Preenche = SUCESSO

    Exit Function

Erro_GridChequeN_Preenche:

    GridChequeN_Preenche = gErr

    Select Case gErr

        Case 103970

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143629)

    End Select

    Exit Function

End Function

Private Sub Move_Tela_Memoria(objBorderoCheque As ClassBorderoCheque)

On Error GoTo Erro_Move_Tela_Memoria

    objBorderoCheque.dtDataBackoffice = DATA_NULA
    objBorderoCheque.dtDataImpressao = DATA_NULA

    'caso contrário preenche com o valo preenchido na tela
    objBorderoCheque.dtDataEnvio = StrParaDate(DataEnvio.Text)


    'preenche os campos de totais
    objBorderoCheque.dValorEspec = StrParaDbl(LabelTotalDesc.Caption)
    objBorderoCheque.dValorNEspec = StrParaDbl(LabelTotalNDesc.Caption)

    'se a conta corrente estiver selecionada
    If ContaCorrente.ListIndex <> -1 Then

        'extrai o codigo e preenche o atributo do obj
        objBorderoCheque.iCodNossaConta = Codigo_Extrai(ContaCorrente.Text)

    End If

    'preenche os campos restantes
    objBorderoCheque.iFilialEmpresa = giFilialEmpresa
    objBorderoCheque.lNumBordero = StrParaLong(Codigo.Text)

    Exit Sub

Erro_Move_Tela_Memoria:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143630)

    End Select

    Exit Sub

End Sub

Private Sub Move_Tela_Memoria_Cheque(colCheque As Collection, colChequeN As Collection)
'função que varre os 2 grids e preenche as coleções de cheques respectivas

Dim iLinha As Integer
Dim objChequePre As ClassChequePre

On Error GoTo Erro_Move_Tela_Memoria_Cheque

    'varre o grid de cheques detalhados
    For iLinha = 1 To objGridCheque.iLinhasExistentes

        'se a coluna selecionado estiver marcada
        If StrParaInt(GridCheques.TextMatrix(iLinha, iGrid_Selecionado_Col)) = MARCADO Then

            'aponta para o cheque correspondente da coleção
            Set objChequePre = gcolCheque.Item(iLinha)

            objChequePre.lNumBorderoLoja = StrParaDbl(Codigo.Text)
            
            'adiciona o cheque à coleção local
            colCheque.Add objChequePre

        End If

    Next

    'varre o grid de cheques não detalhados
    For iLinha = 1 To objGridChequeN.iLinhasExistentes

        'se a coluna selecionado estiver marcada
        If StrParaInt(GridChequeN.TextMatrix(iLinha, iGrid_SelecionadoN_Col)) = MARCADO Then

            'aponta para o cheque correspondente da coleção
            Set objChequePre = gcolChequeN.Item(iLinha)

            objChequePre.lNumBorderoLoja = StrParaDbl(Codigo.Text)
            
            'adiciona o chque à coleção local
            colChequeN.Add objChequePre

        End If

    Next

    Exit Sub

Erro_Move_Tela_Memoria_Cheque:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143631)

    End Select

    Exit Sub

End Sub

Public Function BorderoCheque_Codigo_Automatico(lCodigo As Long) As Long
'função que gera um número de borderô

Dim lErro As Long

On Error GoTo Erro_BorderoCheque_Codigo_Automatico

    'gera um número automático para o bordero
    lErro = CF("Config_ObterAutomatico", "LojaConfig", "COD_PROX_BORDEROCHEQUE", "BorderoCheque", "NumBordero", lCodigo)
    If lErro <> SUCESSO Then gError 103986

    BorderoCheque_Codigo_Automatico = SUCESSO

    Exit Function

Erro_BorderoCheque_Codigo_Automatico:

    BorderoCheque_Codigo_Automatico = gErr

    Select Case gErr

        Case 103986

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143632)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Borderô Cheque"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "BorderoCheque"

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

Private Sub BotaoProxNum_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ActiveControl = Me.ActiveControl

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
