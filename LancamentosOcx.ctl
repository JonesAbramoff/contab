VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl LancamentosOcx 
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11835
   KeyPreview      =   -1  'True
   ScaleHeight     =   6900
   ScaleMode       =   0  'User
   ScaleWidth      =   11840
   Begin VB.TextBox Historico 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   5085
      MaxLength       =   150
      TabIndex        =   9
      Top             =   1785
      Width           =   4725
   End
   Begin VB.CommandButton BotaoHist 
      Caption         =   "Históricos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9900
      TabIndex        =   51
      Top             =   5835
      Width           =   1605
   End
   Begin VB.CommandButton BotaoCcl 
      Caption         =   "Centros de Custo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9915
      TabIndex        =   50
      Top             =   5220
      Width           =   1605
   End
   Begin VB.CommandButton BotaoConta 
      Caption         =   "Plano de Contas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9915
      TabIndex        =   49
      Top             =   4605
      Width           =   1605
   End
   Begin VB.CheckBox Gerencial 
      Height          =   210
      Left            =   5205
      TabIndex        =   48
      Tag             =   "1"
      Top             =   2385
      Width           =   870
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   6675
      Picture         =   "LancamentosOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Numeração Automática"
      Top             =   135
      Width           =   300
   End
   Begin VB.PictureBox Picture3 
      Height          =   525
      Left            =   9360
      ScaleHeight     =   465
      ScaleWidth      =   2310
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   195
      Width           =   2370
      Begin VB.CommandButton BotaoFechar 
         Height          =   330
         Left            =   1875
         Picture         =   "LancamentosOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   330
         Left            =   1395
         Picture         =   "LancamentosOcx.ctx":0268
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Imprimir"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   330
         Left            =   945
         Picture         =   "LancamentosOcx.ctx":079A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   330
         Left            =   60
         Picture         =   "LancamentosOcx.ctx":0CCC
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   330
         Left            =   510
         Picture         =   "LancamentosOcx.ctx":0E26
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   375
      End
   End
   Begin VB.ListBox ListDocAuto 
      Height          =   5520
      Left            =   9240
      TabIndex        =   13
      Top             =   1245
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.ListBox ListHistoricos 
      Height          =   5520
      Left            =   9240
      TabIndex        =   23
      Top             =   1245
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.Frame Frame1 
      Caption         =   "Descrição do Elemento Selecionado"
      Height          =   1050
      Left            =   105
      TabIndex        =   25
      Top             =   4845
      Width           =   6315
      Begin VB.Label CclLabel 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Custo:"
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
         Left            =   240
         TabIndex        =   26
         Top             =   660
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1125
         TabIndex        =   27
         Top             =   300
         Width           =   570
      End
      Begin VB.Label ContaDescricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1845
         TabIndex        =   28
         Top             =   285
         Width           =   3720
      End
      Begin VB.Label CclDescricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1845
         TabIndex        =   29
         Top             =   645
         Visible         =   0   'False
         Width           =   3720
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Documento Automático"
      Height          =   750
      Left            =   105
      TabIndex        =   22
      Top             =   5985
      Width           =   6315
      Begin VB.CommandButton BotaoAplicar 
         Height          =   510
         Left            =   3975
         Picture         =   "LancamentosOcx.ctx":0FB0
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   180
         Width           =   1095
      End
      Begin MSMask.MaskEdBox DocAuto 
         Height          =   285
         Left            =   2190
         TabIndex        =   11
         Top             =   315
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin VB.Label Label6 
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
         Height          =   195
         Left            =   1425
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   30
         Top             =   330
         Width           =   660
      End
   End
   Begin MSMask.MaskEdBox SeqContraPartida 
      Height          =   225
      Left            =   4665
      TabIndex        =   8
      Top             =   1800
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Debito 
      Height          =   225
      Left            =   3465
      TabIndex        =   7
      Top             =   1785
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
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
   Begin MSMask.MaskEdBox Credito 
      Height          =   225
      Left            =   2280
      TabIndex        =   6
      Top             =   1785
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
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
   Begin MSMask.MaskEdBox Ccl 
      Height          =   225
      Left            =   1950
      TabIndex        =   5
      Top             =   2160
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      AllowPrompt     =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Conta 
      Height          =   225
      Left            =   375
      TabIndex        =   4
      Top             =   2190
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      AllowPrompt     =   -1  'True
      MaxLength       =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin MSComctlLib.TreeView TvwContas 
      Height          =   5520
      Left            =   9240
      TabIndex        =   15
      Top             =   1245
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   9737
      _Version        =   393217
      Indentation     =   453
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
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
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   285
      Left            =   2040
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   510
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Documento 
      Height          =   285
      Left            =   5610
      TabIndex        =   1
      Top             =   135
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Lote 
      Height          =   285
      Left            =   3420
      TabIndex        =   0
      Top             =   135
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   285
      Left            =   900
      TabIndex        =   3
      Top             =   510
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComctlLib.TreeView TvwCcls 
      Height          =   5520
      Left            =   9240
      TabIndex        =   14
      Top             =   1245
      Visible         =   0   'False
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   9737
      _Version        =   393217
      Indentation     =   453
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
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
   Begin MSFlexGridLib.MSFlexGrid GridLancamentos 
      Height          =   1860
      Left            =   90
      TabIndex        =   10
      Top             =   1200
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   3281
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin VB.Label LabelDocAuto 
      Caption         =   "Documentos Automáticos"
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
      Left            =   9225
      TabIndex        =   31
      Top             =   1020
      Width           =   2565
   End
   Begin VB.Label LabelCcl 
      Caption         =   "Centros de Custo / Lucro"
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
      Left            =   9240
      TabIndex        =   32
      Top             =   1020
      Width           =   2490
   End
   Begin VB.Label LabelHistoricos 
      Caption         =   "Históricos"
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
      Left            =   9240
      TabIndex        =   33
      Top             =   1020
      Width           =   945
   End
   Begin VB.Label Label5 
      Caption         =   "Período:"
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
      Left            =   4800
      TabIndex        =   34
      Top             =   570
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Origem:"
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
      Height          =   255
      Left            =   165
      TabIndex        =   35
      Top             =   150
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Lote:"
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
      Left            =   2865
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   36
      Top             =   180
      Width           =   450
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   4560
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   37
      Top             =   150
      Width           =   1035
   End
   Begin VB.Label Label4 
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
      Height          =   255
      Left            =   315
      TabIndex        =   38
      Top             =   540
      Width           =   525
   End
   Begin VB.Label LabelTotais 
      Caption         =   "Totais:"
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
      Left            =   1485
      TabIndex        =   39
      Top             =   4290
      Width           =   705
   End
   Begin VB.Label TotalDebito 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3540
      TabIndex        =   40
      Top             =   4275
      Width           =   1155
   End
   Begin VB.Label TotalCredito 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2295
      TabIndex        =   41
      Top             =   4260
      Width           =   1155
   End
   Begin VB.Label Periodo 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5610
      TabIndex        =   42
      Top             =   555
      Width           =   1185
   End
   Begin VB.Label Exercicio 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3420
      TabIndex        =   43
      Top             =   525
      Width           =   1185
   End
   Begin VB.Label Label8 
      Caption         =   "Exercício:"
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
      Left            =   2445
      TabIndex        =   44
      Top             =   540
      Width           =   870
   End
   Begin VB.Label Origem 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   870
      TabIndex        =   45
      Top             =   135
      Width           =   1530
   End
   Begin VB.Label LabelContas 
      Caption         =   "Plano de Contas"
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
      Left            =   9210
      TabIndex        =   46
      Top             =   1020
      Width           =   2340
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Lançamentos"
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
      TabIndex        =   47
      Top             =   990
      Width           =   1140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   11809.99
      Y1              =   930
      Y2              =   930
   End
End
Attribute VB_Name = "LancamentosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iGrid_Conta_Col As Integer
Dim iGrid_Ccl_Col As Integer
Dim iGrid_Debito_Col As Integer
Dim iGrid_Credito_Col As Integer
Dim iGrid_SeqContraPartida_Col As Integer
Dim iGrid_Historico_Col As Integer
Dim iGrid_Gerencial_Col As Integer

Dim objGrid1 As AdmGrid
Dim iAlterado As Integer

Private WithEvents objEventoDocAuto As AdmEvento
Attribute objEventoDocAuto.VB_VarHelpID = -1
Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoLancamento As AdmEvento
Attribute objEventoLancamento.VB_VarHelpID = -1

Private WithEvents objEventoConta As AdmEvento
Attribute objEventoConta.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoHist As AdmEvento
Attribute objEventoHist.VB_VarHelpID = -1


Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lDoc As Long
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo

On Error GoTo Erro_BotaoProxNum_Click

    If Len(Data.ClipText) = 0 Then Error 55713

    'Obtém Periodo e Exercicio correspondentes à data
    dtData = CDate(Data.Text)

    lErro = CF("Periodo_Le", dtData, objPeriodo)
    If lErro <> SUCESSO Then Error 5901

    'Mostra número do próximo voucher(documento) disponível
    lErro = CF("Voucher_Automatico", giFilialEmpresa, objPeriodo.iExercicio, objPeriodo.iPeriodo, MODULO_CONTABILIDADE, lDoc)
    If lErro <> SUCESSO Then Error 5833

    Documento.Text = CStr(lDoc)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 5833, 5901
        
        Case 55713
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PROXNUM_DATA_NAO_PREENCHIDA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162242)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoAplicar_Click()

Dim lErro As Long
Dim lDoc As Long

On Error GoTo Erro_BotaoAplicar_click

    If Len(DocAuto.ClipText) = 0 Then Error 11352
    
    lDoc = CLng(DocAuto.ClipText)
    
    lErro = Traz_DocAuto_Tela(lDoc)
    If lErro <> SUCESSO Then Error 11353

    Exit Sub

Erro_BotaoAplicar_click:

    Select Case Err

        Case 11352
        
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO", Err)
        
        Case 11353
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162243)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoImprimir_Click()
'imprime um relatorio com os dados que estao na tela

Dim lErro As Long, objRelTela As New ClassRelTela, iIndice1 As Integer
Dim colTemp As Collection, objLancamento As New ClassLancamento_Detalhe
Dim colLancamento_Detalhe As New Collection
Dim sOrigem As String, sExercicio As String, sPeriodo As String
Dim lDoc As Long, dtData As Date, lLote As Long
Dim sNomeTsk As String
Dim iIndice As Integer

On Error GoTo Erro_BotaoImprimir_Click
    
    lErro = objRelTela.Iniciar("REL_LANC_CTB")
    If lErro <> SUCESSO Then gError 41582
    
    If giTipoVersao = VERSAO_LIGHT Then
        sNomeTsk = "LancCTBL"
    ElseIf giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        sNomeTsk = "Lanc_CCL"
    Else
        sNomeTsk = "Lanc_CTB"
    End If

    'obter dados comuns a todas as linhas do grid
    sOrigem = gobjColOrigem.Origem(Origem.Caption)
    sExercicio = Exercicio.Caption
    sPeriodo = Periodo.Caption
    
    lDoc = StrParaLong(Documento.ClipText)
    dtData = StrParaDate(Data.Text)
    lLote = StrParaLong(Lote.ClipText)
    
    lErro = Grid_Lancamento_Detalhe(colLancamento_Detalhe)
    If lErro <> SUCESSO Then gError 41583
    
    For iIndice1 = 1 To colLancamento_Detalhe.Count
    
        Set objLancamento = colLancamento_Detalhe.Item(iIndice1)
        
        Set colTemp = New Collection
        
        'incluir os valores na mesma ordem da tabela RelTelaCampos no dicdados
        
        Call colTemp.Add(sOrigem)
        Call colTemp.Add(sExercicio)
        Call colTemp.Add(sPeriodo)
        Call colTemp.Add(lDoc)
        Call colTemp.Add(iIndice1)
        Call colTemp.Add(lLote)
        Call colTemp.Add(dtData)
        Call colTemp.Add(objLancamento.sConta)
        Call colTemp.Add(objLancamento.sCcl)
        Call colTemp.Add(objLancamento.sHistorico)
        Call colTemp.Add(objLancamento.dValor)
        Call colTemp.Add(objLancamento.iSeqContraPartida)

        lErro = objRelTela.IncluirRegistro(colTemp)
        If lErro <> SUCESSO Then gError 41584
    
    Next
    
    lErro = objRelTela.ExecutarRel(sNomeTsk)
    If lErro <> SUCESSO Then gError 41585
    
    Exit Sub
    
Erro_BotaoImprimir_Click:

    Select Case gErr
          
        Case 41582, 41583, 41584, 41585
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162244)
     
    End Select

    Exit Sub

End Sub

Private Sub Ccl_Change()

        iAlterado = REGISTRO_ALTERADO
        
End Sub

Private Sub Conta_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Conta_GotFocus()

Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_Conta_GotFocus

    Call Grid_Campo_Recebe_Foco(objGrid1)
    
'    TvwContas.Visible = True
'    LabelContas.Visible = True
'    TvwCcls.Visible = False
'    LabelCcl.Visible = False
'    ListDocAuto.Visible = False
'    LabelDocAuto.Visible = False
'    ListHistoricos.Visible = False
'    LabelHistoricos.Visible = False
    
    'Comentado por Wagner - Dá erro em Loop
'    sConta = GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Conta_Col)
'
'    If Len(sConta) > 0 Then
'
'        lErro = Conta_Exibe_Descricao(sConta)
'        If lErro <> SUCESSO Then Error 5852
'
'    Else
'
'        ContaDescricao = ""
'
'    End If
    
    
    Exit Sub
    
Erro_Conta_GotFocus:

    Select Case Err
    
        Case 5852
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162245)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub Conta_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub Conta_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Conta
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Ccl_GotFocus()

Dim sCcl As String
Dim lErro As Long

On Error GoTo Erro_Ccl_GotFocus

    Call Grid_Campo_Recebe_Foco(objGrid1)
    
'    TvwCcls.Visible = True
'    LabelCcl.Visible = True
'    TvwContas.Visible = False
'    LabelContas.Visible = False
'    ListDocAuto.Visible = False
'    LabelDocAuto.Visible = False
'    ListHistoricos.Visible = False
'    LabelHistoricos.Visible = False
   
'    sCcl = GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col)
'
'    'Coloca descricao de Ccl no panel
'    If Len(sCcl) > 0 Then
'
'        lErro = Ccl_Exibe_Descricao(sCcl)
'        If lErro <> SUCESSO Then Error 5863
'
'    Else
'
'        CclDescricao = ""
'
'    End If
    
    Exit Sub
    
Erro_Ccl_GotFocus:

    Select Case Err
    
        Case 5863
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162246)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Private Sub DocAuto_GotFocus()
    
'    ListDocAuto.Visible = True
'    LabelDocAuto.Visible = True
'    TvwCcls.Visible = False
'    LabelCcl.Visible = False
'    TvwContas.Visible = False
'    LabelContas.Visible = False
'    ListHistoricos.Visible = False
'    LabelHistoricos.Visible = False
    
    Call MaskEdBox_TrataGotFocus(DocAuto, iAlterado)

End Sub

Private Sub Ccl_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Ccl_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Ccl
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Credito_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Credito_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Credito_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub Credito_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Credito
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Documento_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Documento, iAlterado)

End Sub

Private Sub Lote_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Lote, iAlterado)

End Sub

Private Sub SeqContraPartida_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub SeqContraPartida_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub SeqContraPartida_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub SeqContraPartida_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = SeqContraPartida
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Debito_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Debito_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Debito_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub Debito_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Debito
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub DocAuto_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Documento_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Exercicio_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
Dim lErro As Long
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
        
    Set objEventoLote = Nothing
    Set objEventoLancamento = Nothing
    Set objEventoDocAuto = Nothing
    
    Set objEventoConta = Nothing
    Set objEventoCcl = Nothing
    Set objEventoHist = Nothing
    
    Set objGrid1 = Nothing
    
End Sub

Private Sub GridLancamentos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Faz com que apareca um PopupMenu o botao direito do mouse acionado sobre o grid

    'Verifica se foi o botao direito do mouse que foi pressionado
    If Button = vbRightButton Then
        Set PopupMenusCTB.objTela = Me
        PopupMenu PopupMenusCTB.MenuGrid
        Set PopupMenusCTB.objTela = Nothing
    End If

End Sub

Private Sub Historico_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Historico_GotFocus()
    
Dim iPos As Integer
    
    Call Grid_Campo_Recebe_Foco(objGrid1)
    
'    ListHistoricos.Visible = True
'    LabelHistoricos.Visible = True
'    TvwCcls.Visible = False
'    LabelCcl.Visible = False
'    TvwContas.Visible = False
'    LabelContas.Visible = False
'    ListDocAuto.Visible = False
'    LabelDocAuto.Visible = False
    
    If Len(Historico.Text) > 0 Then
        iPos = InStr(Historico.Text, CARACTER_HISTORICO_PARAM)
        If iPos > 0 Then
            Historico.SelStart = iPos - 1
            Historico.SelLength = 1
        End If
    End If
    
End Sub

Private Sub Historico_KeyPress(KeyAscii As Integer)

Dim iInicio As Integer
Dim iPos As Integer
Dim sValor As String
Dim lErro As Long
Dim objHistPadrao As New ClassHistPadrao
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Historico_KeyPress

    'se digitou ENTER
    If KeyAscii = vbKeyReturn Then
        
        If Len(Historico.Text) > 0 Then
        
            If left(Historico.Text, 1) = CARACTER_HISTPADRAO Then
            
                sValor = Trim(Mid(Historico.Text, 2))
                
                lErro = Valor_Inteiro_Critica(sValor)
                If lErro <> SUCESSO Then Error 44073
                
                objHistPadrao.iHistPadrao = CInt(sValor)
                        
                lErro = CF("HistPadrao_Le", objHistPadrao)
                If lErro <> SUCESSO And lErro <> 5446 Then Error 44074
                
                If lErro = 5446 Then Error 44075
        
                Historico.Text = objHistPadrao.sDescHistPadrao
                Historico.SelStart = 0
                
            End If
    
            If Historico.SelText = CARACTER_HISTORICO_PARAM Then
                iInicio = Historico.SelStart + 2
            Else
                iInicio = Historico.SelStart
            End If
        
            If iInicio = 0 Then iInicio = 1
        
            iPos = InStr(iInicio, Historico.Text, CARACTER_HISTORICO_PARAM)
            If iPos > 0 Then
                Historico.SelStart = iPos - 1
                Historico.SelLength = 1
                Exit Sub
            End If
        End If
    End If

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
    Exit Sub
    
Erro_Historico_KeyPress:

    Select Case Err
    
        Case 44073
            objGrid1.iExecutaSaidaCelula = 0
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_HISTPADRAO_INVALIDO", Err, sValor)
            objGrid1.iExecutaSaidaCelula = 1
        
        Case 44074

        Case 44075
            objGrid1.iExecutaSaidaCelula = 0
            
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_HISTPADRAO_INEXISTENTE", objHistPadrao.iHistPadrao)

            If vbMsgRes = vbYes Then
            
                Call Chama_Tela("HistoricoPadrao", objHistPadrao)
            
            Else
                Historico.SetFocus
            End If
            
            objGrid1.iExecutaSaidaCelula = 1
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162247)
    
    End Select

    Exit Sub
    
End Sub

Private Sub Historico_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Historico
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub GridLancamentos_Click()

Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If
    
End Sub

Private Sub GridLancamentos_GotFocus()
    
    Call Grid_Recebe_Foco(objGrid1)

End Sub

Private Sub GridLancamentos_EnterCell()
        
    Call Grid_Entrada_Celula(objGrid1, iAlterado)
    
End Sub

Private Sub GridLancamentos_LeaveCell()
    
    Call Saida_Celula(objGrid1)
    
    objGrid1.iLinhaAntiga = GridLancamentos.Row 'Inserido por Wagner
    
End Sub

Private Sub GridLancamentos_KeyDown(KeyCode As Integer, Shift As Integer)

Dim dColunaSoma As Double
Dim lErro As Long

On Error GoTo Erro_GridLancamentos_KeyDown

    lErro = Grid_Trata_Tecla1(KeyCode, objGrid1)
    If lErro <> SUCESSO Then Error 44065
    
    Call Trata_SeqContraPartida(GridLancamentos.Row)
    dColunaSoma = GridColuna_Soma(iGrid_Debito_Col)
    TotalDebito = Format(dColunaSoma, "Standard")
    dColunaSoma = GridColuna_Soma(iGrid_Credito_Col)
    TotalCredito = Format(dColunaSoma, "Standard")
    
    Exit Sub
    
Erro_GridLancamentos_KeyDown:

    Select Case Err
    
        Case 44065
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162248)
    
    End Select

    Exit Sub

End Sub

Private Sub GridLancamentos_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridLancamentos_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGrid1)

End Sub

Private Sub GridLancamentos_RowColChange()

Dim iLinhaAnterior As Integer

    iLinhaAnterior = objGrid1.iLinhaAntiga

    Call Grid_RowColChange(objGrid1)
    
    If iLinhaAnterior <> GridLancamentos.Row Then Call Exibe_Dados    'Inserido por Wagner
       
End Sub

Private Sub GridLancamentos_Scroll()

    Call Grid_Scroll(objGrid1)
    
End Sub

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoLote = New AdmEvento
    Set objEventoLancamento = New AdmEvento
    Set objEventoDocAuto = New AdmEvento
            
    Set objEventoConta = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoHist = New AdmEvento
            
            
    TvwContas.Visible = False
    LabelContas.Visible = False
    TvwCcls.Visible = False
    LabelCcl.Visible = False
    ListDocAuto.Visible = False
    LabelDocAuto.Visible = False
    ListHistoricos.Visible = False
    LabelHistoricos.Visible = False
            
            
    Set objGrid1 = New AdmGrid
    
    lErro = Inicializa_Grid_Lancamento(objGrid1)
    If lErro <> SUCESSO Then Error 5899

'    'Inicializa a Lista de Plano de Contas
'    lErro = CF("Carga_Arvore_Conta", TvwContas.Nodes)
'    If lErro <> SUCESSO Then Error 5860

    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
    
'        'Inicializa a Lista de Centros de Custo
'        lErro = Carga_Arvore_Ccl(TvwCcls.Nodes)
'        If lErro <> SUCESSO Then Error 5861
        
    End If
    
'    'Inicializa a Lista de Historicos
'    lErro = Carga_Lista_Historico()
'    If lErro <> SUCESSO Then Error 5892

    'Inicializa a Lista de Documentos Automaticos
    lErro = Carga_Lista_DocAuto()
    If lErro <> SUCESSO Then Error 11348

    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        CclLabel.Visible = True
        CclDescricao.Visible = True
    End If
    
    Origem.Caption = "Contabilidade"
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case 5860, 5861, 5892, 5899, 11348
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162249)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objLancamento_Detalhe As ClassLancamento_Detalhe) As Long

Dim lErro As Long
Dim lDoc As Long
Dim objDocAuto As ClassDocAuto
Dim objLancamento_Cabecalho As ClassLancamento_Cabecalho

On Error GoTo Erro_Trata_Parametros

    'Se há um documento(voucher) passado como parametro, exibir seus dados
    If Not (objLancamento_Detalhe Is Nothing) Then
    
        Set objLancamento_Cabecalho = New ClassLancamento_Cabecalho
    
        objLancamento_Cabecalho.iFilialEmpresa = objLancamento_Detalhe.iFilialEmpresa
        objLancamento_Cabecalho.sOrigem = objLancamento_Detalhe.sOrigem
        objLancamento_Cabecalho.iExercicio = objLancamento_Detalhe.iExercicio
        objLancamento_Cabecalho.iPeriodoLan = objLancamento_Detalhe.iPeriodoLan
        objLancamento_Cabecalho.lDoc = objLancamento_Detalhe.lDoc
    
        lErro = Traz_Doc_Tela(objLancamento_Cabecalho)
        If lErro <> SUCESSO And lErro <> 5843 Then Error 5849
        
    Else
        
        Call Limpa_Tela_Lancamentos
        
        lErro = Traz_Cabecalho_Tela()
        If lErro <> SUCESSO Then Error 5848
      
        iAlterado = 0
        
    End If
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case 5848, 5849
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162250)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Private Function Carga_Arvore_Ccl(colNodes As Nodes) As Long
'move os dados de centro de custo/lucro do banco de dados para a arvore colNodes.

Dim objNode As Node
Dim colCcl As New Collection
Dim objCcl As ClassCcl
Dim lErro As Long
Dim sCclMascarado As String
Dim sCcl As String
Dim sCclPai As String
    
On Error GoTo Erro_Carga_Arvore_Ccl
    
    lErro = CF("Ccl_Le_Todos", colCcl)
    If lErro <> SUCESSO Then Error 10496
    
    For Each objCcl In colCcl
        
        sCclMascarado = String(STRING_CCL, 0)

        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then Error 10497

        If objCcl.iTipoCcl = CCL_ANALITICA Then
            sCcl = "A" & objCcl.sCcl
        Else
            sCcl = "S" & objCcl.sCcl
        End If

        sCclPai = String(STRING_CCL, 0)
        
        'retorna o centro de custo/lucro "pai" do centro de custo/lucro em questão, se houver
        lErro = Mascara_RetornaCclPai(objCcl.sCcl, sCclPai)
        If lErro <> SUCESSO Then Error 10498
        
        'se o centro de custo/lucro possui um centro de custo/lucro "pai"
        If Len(Trim(sCclPai)) > 0 Then

            sCclPai = "S" & sCclPai
            
            Set objNode = colNodes.Add(colNodes.Item(sCclPai), tvwChild, sCcl)

        Else
            'se o centro de custo/lucro não possui centro de custo/lucro "pai"
            Set objNode = colNodes.Add(, tvwLast, sCcl)
            
        End If
        
        objNode.Text = sCclMascarado & SEPARADOR & objCcl.sDescCcl
        
    Next
    
    Carga_Arvore_Ccl = SUCESSO

    Exit Function

Erro_Carga_Arvore_Ccl:

    Carga_Arvore_Ccl = Err

    Select Case Err

        Case 10496
        
        Case 10497
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objCcl.sCcl)

        Case 10498
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCclPai", Err, objCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162251)

    End Select
    
    Exit Function

End Function

Private Function Carga_Lista_Historico() As Long
'move os dados de historico padrão do banco de dados para a arvore colNodes.

Dim colHistPadrao As New Collection
Dim objHistPadrao As ClassHistPadrao
Dim lErro As Long
    
On Error GoTo Erro_Carga_Lista_Historico
    
    lErro = CF("HistPadrao_Le_Todos", colHistPadrao)
    If lErro <> SUCESSO Then Error 5893
    
    For Each objHistPadrao In colHistPadrao
        
        ListHistoricos.AddItem CStr(objHistPadrao.iHistPadrao) & SEPARADOR & objHistPadrao.sDescHistPadrao
        ListHistoricos.ItemData(ListHistoricos.NewIndex) = objHistPadrao.iHistPadrao
        
    Next
    
    Carga_Lista_Historico = SUCESSO

    Exit Function

Erro_Carga_Lista_Historico:

    Carga_Lista_Historico = Err

    Select Case Err

        Case 5893

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162252)

    End Select
    
    Exit Function

End Function

Private Function Carga_Lista_DocAuto() As Long
'move os dados do plano de contas do banco de dados para a arvore colNodes.

Dim colDocAuto As New Collection
Dim objDocAuto As ClassDocAuto
Dim lErro As Long
    
On Error GoTo Erro_Carga_Lista_DocAuto
    
    'leitura das contas no BD
    lErro = CF("DocAuto_Le_Todos", colDocAuto)
    If lErro <> SUCESSO Then Error 5911
    
    For Each objDocAuto In colDocAuto
        
        ListDocAuto.AddItem CStr(objDocAuto.lDoc) & SEPARADOR & objDocAuto.sDescricao
                
    Next
    
    Carga_Lista_DocAuto = SUCESSO

    Exit Function

Erro_Carga_Lista_DocAuto:

    Carga_Lista_DocAuto = Err

    Select Case Err

        Case 5911
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162253)

    End Select
    
    Exit Function

End Function

Private Function Traz_Cabecalho_Tela() As Long

Dim sDescricao As String
Dim iPeriodoDoc As Integer
Dim iExercicioDoc As Integer
Dim iIndice As Integer
Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim lErro As Long

On Error GoTo Erro_Traz_Cabecalho_Tela

    'Inicializa Data
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    'Coloca o periodo relativo a data na tela
    lErro = CF("Periodo_Le", gdtDataAtual, objPeriodo)
    If lErro <> SUCESSO Then Error 6091
    
    Periodo.Caption = objPeriodo.sNomeExterno
    
    lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
    If lErro <> SUCESSO And lErro <> 10083 Then Error 5821
    
    If lErro = 10083 Then Error 10084
    
    Exercicio.Caption = objExercicio.sNomeExterno

    Traz_Cabecalho_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Cabecalho_Tela:

    Traz_Cabecalho_Tela = Err

    Select Case Err
    
        Case 5821, 6091
            
        Case 10084
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162254)
    
    End Select
    
    Exit Function

End Function

Private Function Traz_Doc_Tela(objDoc As ClassLancamento_Cabecalho) As Long
'traz os dados do voucher do banco de dados para a tela

Dim lErro As Long
Dim colLancamentos As New Collection
Dim objLanc As ClassLancamento_Detalhe
Dim iLinha As Integer
Dim sContaMascarada As String
Dim sCclMascarado As String
Dim sDescricao As String
Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim iIndice As Integer
Dim dAcumulador As Double
Dim iLinhaAtual As Integer
Dim iColunaAtual As Integer
Dim objProduto As New ClassProduto
Dim objEstoqueMes As New ClassEstoqueMes
Dim iQuantidade As Integer
Dim iFilialEmpresaSalva As Integer
Dim iAchou As Integer
Dim iIndice1 As Integer

On Error GoTo Erro_Traz_Doc_Tela
    
    Call Limpa_Tela_Lancamentos
    
    
    iLinhaAtual = GridLancamentos.Row
    iColunaAtual = GridLancamentos.Col
    
    Origem.Caption = gobjColOrigem.Descricao(objDoc.sOrigem)
    
    'Origem só pode ser CTB ou FLH
    If gobjColOrigem.Origem(Origem.Caption) <> "CTB" And gobjColOrigem.Origem(Origem.Caption) <> "FLH" Then
    
        objGrid1.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
        objGrid1.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
        
    End If
    
    Documento.Text = CStr(objDoc.lDoc)

    iFilialEmpresaSalva = objDoc.iFilialEmpresa

    Do While objDoc.iFilialEmpresa > 0 And objDoc.iFilialEmpresa < 100

        'Lê os lançamentos contidos no documento em questão
        lErro = CF("LanPendente_Le_Doc", objDoc, colLancamentos)
        If lErro <> SUCESSO And lErro <> 5842 Then Error 5841

        If lErro = SUCESSO Then iAchou = 1

        If giContabGerencial = 0 Then Exit Do

        objDoc.iFilialEmpresa = objDoc.iFilialEmpresa - giFilialAuxiliar

    Loop

    objDoc.iFilialEmpresa = iFilialEmpresaSalva
    
    'se não encontrou o documento
    If iAchou = 0 Then gError 5843
    
    For iIndice = colLancamentos.Count To 1 Step -1
        
        For iIndice1 = iIndice - 1 To 1 Step -1
        
            If colLancamentos(iIndice).iSeq = colLancamentos(iIndice1).iSeq Then
                colLancamentos.Remove (iIndice)
                Exit For
            End If
            
        Next
    
    Next
    
    Set objLanc = colLancamentos.Item(1)
    
    Lote.Text = CStr(objLanc.iLote)
    
    'Inicializa Data
    Data.Text = Format(objLanc.dtData, "dd/mm/yy")
    
    'Coloca o periodo relativo a data na tela
    lErro = CF("Periodo_Le", objLanc.dtData, objPeriodo)
    If lErro <> SUCESSO Then gError 5846
    
    Periodo.Caption = objPeriodo.sNomeExterno
    
    'Coloca o exercicio na tela
    lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
    If lErro <> SUCESSO And lErro <> 10083 Then gError 5847
    
    'Se o exercicio não está cadastrado
    If lErro = 10083 Then gError 10085
    
    Exercicio.Caption = objExercicio.sNomeExterno
    
    If colLancamentos.Count > MAX_LANCAMENTOS_POR_DOC_CTB + 1 Then gError 197911
    
    If colLancamentos.Count >= objGrid1.objGrid.Rows Then
        Call Refaz_Grid(objGrid1, colLancamentos.Count)
    End If
    
    'move os dados para a tela
    For Each objLanc In colLancamentos
    
        objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
    
        If Len(objLanc.sConta) > 0 Then
    
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objLanc.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 5844
            
            Conta.PromptInclude = False
            Conta.Text = sContaMascarada
            Conta.PromptInclude = True
            
            'coloca a conta na tela
            GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Conta_Col) = Conta.Text
            
        Else
        
            GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Conta_Col) = ""
            
        End If
        
        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        
            If Len(objLanc.sCcl) > 0 Then
        
                'mascara o centro de custo
                sCclMascarado = String(STRING_CCL, 0)
               
               lErro = Mascara_MascararCcl(objLanc.sCcl, sCclMascarado)
                If lErro <> SUCESSO Then gError 55505
                
                Ccl.PromptInclude = False
                Ccl.Text = sCclMascarado
                Ccl.PromptInclude = True
                
                'coloca o centro de custo na tela
                GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Ccl_Col) = Ccl.Text
                
            Else
            
                GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Ccl_Col) = ""
            
            End If
            
        End If
        
        If Len(Trim(objLanc.sProduto)) > 0 Then
                
            objProduto.sCodigo = objLanc.sProduto
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO Then gError 83635
        
            'se a contabilização está associada a um produto apropriado pelo custo de produção
            If objProduto.iApropriacaoCusto = APROPR_CUSTO_REAL Or objProduto.iApropriacaoCusto = APROPR_CUSTO_MEDIO_PRODUCAO Then
                
                objEstoqueMes.iFilialEmpresa = objLanc.iFilialEmpresa
                objEstoqueMes.iAno = Year(objLanc.dtDataEstoque)
                objEstoqueMes.iMes = Month(objLanc.dtDataEstoque)
                
                'verifica se o mes relativo a data do movimento em questão não está fechada
                lErro = CF("EstoqueMes_Le", objEstoqueMes)
                If lErro <> SUCESSO And lErro <> 36513 Then gError 83636
                
                'se o estoquemes não estiver cadastrado ==> erro
                If lErro = 36513 Then gError 83637
                
                'se o custo de produção ainda não foi apurado ==> o valor do lançamento é quantidade
                If objEstoqueMes.iCustoProdApurado = CUSTO_NAO_APURADO Then iQuantidade = 1
                
            End If
        
        End If
        
        'coloca o valor na tela
        If objLanc.dValor > 0 Then
            GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Credito_Col) = Format(objLanc.dValor, "Standard")
            
            'se o valor for quantidade ==> coloca em bold
            If iQuantidade = 1 Then
            
                GridLancamentos.Row = objLanc.iSeq
                GridLancamentos.Col = iGrid_Credito_Col
                GridLancamentos.CellFontBold = True
                        
            End If
        Else
            GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Debito_Col) = Format(-objLanc.dValor, "Standard")
            
            'se o valor for quantidade ==> coloca em bold
            If iQuantidade = 1 Then
            
                GridLancamentos.Row = objLanc.iSeq
                GridLancamentos.Col = iGrid_Debito_Col
                GridLancamentos.CellFontBold = True
            End If
        End If
            
            
        If objLanc.iSeqContraPartida <> 0 Then GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_SeqContraPartida_Col) = CStr(objLanc.iSeqContraPartida)
            
        'coloca o historico na tela
        GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Historico_Col) = objLanc.sHistorico
            
        If giContabGerencial = 1 Then GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Gerencial_Col) = CStr(objLanc.iGerencial)
            
    Next
    
    dAcumulador = GridColuna_Soma(iGrid_Credito_Col)
    TotalCredito.Caption = Format(dAcumulador, "Standard")
    
    dAcumulador = GridColuna_Soma(iGrid_Debito_Col)
    TotalDebito.Caption = Format(dAcumulador, "Standard")
    
    GridLancamentos.Row = iLinhaAtual
    GridLancamentos.Col = iColunaAtual
    
    Call Grid_Refresh_Checkbox(objGrid1)
    
    iAlterado = 0
    
    Traz_Doc_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Doc_Tela:

    Traz_Doc_Tela = gErr

    Select Case gErr
    
        Case 5841, 5843, 5846, 5847, 83635, 83636
        
        Case 5844
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objLanc.sConta)
        
        Case 10085
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", gErr, objPeriodo.iExercicio)
        
        Case 55505
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objLanc.sCcl)
        
        Case 83637
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_INEXISTENTE", gErr, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)
        
        Case 197911
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_LANC_MAIOR_LIMITE", gErr, colLancamentos.Count, MAX_LANCAMENTOS_POR_DOC_CTB)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162255)
        
    End Select
    
    iAlterado = 0
    
    Exit Function
        
End Function

Private Function Inicializa_Grid_Lancamento(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_Lancamento
    
    'tela em questão
    Set objGrid1.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Conta")
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then objGridInt.colColuna.Add ("CCusto")
    objGridInt.colColuna.Add ("Débito")
    objGridInt.colColuna.Add ("Crédito")
    objGridInt.colColuna.Add ("C.P.")
    objGridInt.colColuna.Add ("Histórico")
    If giContabGerencial = 1 Then objGridInt.colColuna.Add ("Status")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (Conta.Name)
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (Debito.Name)
    objGridInt.colCampo.Add (Credito.Name)
    objGridInt.colCampo.Add (SeqContraPartida.Name)
    objGridInt.colCampo.Add (Historico.Name)
    If giContabGerencial = 1 Then objGridInt.colCampo.Add (Gerencial.Name)
    
    'indica onde estão situadas as colunas do grid
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        iGrid_Conta_Col = 1
        iGrid_Ccl_Col = 2
        iGrid_Debito_Col = 3
        iGrid_Credito_Col = 4
        iGrid_SeqContraPartida_Col = 5
        iGrid_Historico_Col = 6
    Else
        iGrid_Conta_Col = 1
        '999 indica que não está sendo usado
        iGrid_Ccl_Col = 999
        iGrid_Debito_Col = 2
        iGrid_Credito_Col = 3
        iGrid_SeqContraPartida_Col = 4
        iGrid_Historico_Col = 5
        Ccl.Visible = False
    End If
    
    If giContabGerencial = 1 Then
        iGrid_Gerencial_Col = iGrid_Historico_Col + 1
    Else
        Gerencial.Visible = False
    End If
    
    lErro = Inicializa_Mascaras()
    If lErro <> SUCESSO Then Error 5724

    objGridInt.objGrid = GridLancamentos
    
    'todas as linhas do grid
'    objGridInt.objGrid.Rows = MAX_LANCAMENTOS_POR_DOC_CTB + 1
'    objGridInt.objGrid.Rows = 100
    objGridInt.objGrid.Rows = 400
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 12 '7
        
    GridLancamentos.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGridInt)

    'Posiciona os painéis totalizadores
    TotalDebito.top = GridLancamentos.top + GridLancamentos.Height
    TotalDebito.left = GridLancamentos.left
    For iIndice = 0 To iGrid_Debito_Col - 1
        TotalDebito.left = TotalDebito.left + GridLancamentos.ColWidth(iIndice) + GridLancamentos.GridLineWidth + 20
    Next
    
    TotalDebito.Width = GridLancamentos.ColWidth(iGrid_Debito_Col)
    
    TotalCredito.top = TotalDebito.top
    TotalCredito.left = TotalDebito.left + TotalDebito.Width + GridLancamentos.GridLineWidth
    TotalCredito.Width = GridLancamentos.ColWidth(iGrid_Credito_Col)
    
    LabelTotais.top = TotalCredito.top + (TotalDebito.Height - LabelTotais.Height) / 2
    LabelTotais.left = TotalDebito.left - LabelTotais.Width

    Inicializa_Grid_Lancamento = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_Lancamento:

    Inicializa_Grid_Lancamento = Err
    
    Select Case Err
    
        Case 5724
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162256)
        
    End Select

    Exit Function
        
End Function

Private Function Inicializa_Mascaras() As Long
'inicializa as mascaras de conta e centro de custo

Dim sMascaraConta As String
Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascaras

    'Inicializa a máscara de Conta
    sMascaraConta = String(STRING_CONTA, 0)
    
    'le a mascara das contas
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then Error 5694
    
    Conta.Mask = sMascaraConta
    
    'Se usa centro de custo/lucro ==> inicializa mascara de centro de custo/lucro
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
    
        sMascaraCcl = String(STRING_CCL, 0)

        'le a mascara dos centros de custo/lucro
        lErro = MascaraCcl(sMascaraCcl)
        If lErro <> SUCESSO Then Error 5695

        Ccl.Mask = sMascaraCcl
        
    End If
    
    Inicializa_Mascaras = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascaras:

    Inicializa_Mascaras = Err
    
    Select Case Err
    
        Case 5694, 5695
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162257)
        
    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iUltimaLinha As Integer
Dim ColRateioOn As New Collection

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then
    
        Select Case GridLancamentos.Col
    
            Case iGrid_Conta_Col
            
                lErro = Saida_Celula_Conta(objGridInt)
                If lErro <> SUCESSO Then gError 5865
                
            Case iGrid_Ccl_Col
            
                lErro = Saida_Celula_Ccl(objGridInt)
                If lErro <> SUCESSO Then gError 5867
                
            Case iGrid_Credito_Col
            
                lErro = Saida_Celula_Credito(objGridInt)
                If lErro <> SUCESSO Then gError 5879
                
            Case iGrid_Debito_Col
            
                lErro = Saida_Celula_Debito(objGridInt)
                If lErro <> SUCESSO Then gError 5880

            Case iGrid_SeqContraPartida_Col
            
                lErro = Saida_Celula_SeqContraPartida(objGridInt)
                If lErro <> SUCESSO Then gError 20618

            Case iGrid_Historico_Col
            
                lErro = Saida_Celula_Historico(objGridInt)
                If lErro <> SUCESSO Then gError 5898
               
            Case iGrid_Gerencial_Col
            
                lErro = Saida_Celula_Gerencial(objGridInt)
                If lErro <> SUCESSO Then gError 188074

        End Select
    
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 5715
        
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula:

    Saida_Celula = gErr
    
    Select Case gErr
    
        Case 5715
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 5865, 5867, 5879, 5880, 5898, 20618, 188074
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162258)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Conta(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim lPosicaoSeparador As Long
Dim sContaFormatada As String
Dim sCclFormatada As String
Dim iContaPreenchida As Integer
Dim iCclPreenchida As Integer
Dim objContaCcl As New ClassContaCcl
Dim sCcl As String
Dim vbMsgRes As VbMsgBoxResult
Dim objPlanoConta As New ClassPlanoConta
Dim sContaMascarada As String

On Error GoTo Erro_Saida_Celula_Conta

    Set objGridInt.objControle = Conta
    
    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica", Conta.Text, Conta.ClipText, objPlanoConta)
    If lErro <> SUCESSO And lErro <> 44033 And lErro <> 44037 Then Error 20603
    
    'se é uma conta simples, coloca a conta normal no lugar da conta simples
    If lErro = SUCESSO Then
    
        sContaFormatada = objPlanoConta.sConta
        
        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)
        
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 20639
        
        Conta.PromptInclude = False
        Conta.Text = sContaMascarada
        Conta.PromptInclude = True
        
        'Se a Conta possui um Histórico Padrão "default" coloca na tela
        If Len(Trim(GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Historico_Col))) = 0 And objPlanoConta.iHistPadrao <> 0 Then
                        
            For iIndice = 0 To ListHistoricos.ListCount - 1
                If ListHistoricos.ItemData(iIndice) = objPlanoConta.iHistPadrao Then
                    ListHistoricos.ListIndex = iIndice
                    lPosicaoSeparador = InStr(ListHistoricos.Text, SEPARADOR)
                    GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Historico_Col) = Mid(ListHistoricos.Text, lPosicaoSeparador + 1)
                    Exit For
                End If
            Next
            
        End If

    'se não encontrou a conta simples
    ElseIf lErro = 44033 Or lErro = 44037 Then
    
        'testa a conta no seu formato normal
        'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", Conta.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 5696
                
        'conta não cadastrada
        If lErro = 5700 Then Error 5806
        
        'Se a Conta possui um Histórico Padrão "default" coloca na tela
        If Len(Trim(GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Historico_Col))) = 0 And objPlanoConta.iHistPadrao <> 0 Then
                        
            For iIndice = 0 To ListHistoricos.ListCount - 1
                If ListHistoricos.ItemData(iIndice) = objPlanoConta.iHistPadrao Then
                    ListHistoricos.ListIndex = iIndice
                    lPosicaoSeparador = InStr(ListHistoricos.Text, SEPARADOR)
                    GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Historico_Col) = Mid(ListHistoricos.Text, lPosicaoSeparador + 1)
                    Exit For
                End If
            Next
            
        End If

    End If
    
    'se a conta foi preenchida
    If Len(Conta.ClipText) > 0 Then
    
        'se utiliza centro de custo extra-contabil
        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        
            'se o centro de custo foi preenchido
            If Len(GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Ccl_Col)) > 0 Then
            
                'verifica se a associação da conta com o centro de custo está cadastrado
                sCcl = GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Ccl_Col)
        
                lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
                If lErro <> SUCESSO Then Error 5873
        
                objContaCcl.sConta = sContaFormatada
                objContaCcl.sCcl = sCclFormatada
        
                lErro = CF("ContaCcl_Le", objContaCcl)
                If lErro <> SUCESSO And lErro <> 5871 Then Error 5874
        
                'associação Conta x Centro de Custo/Lucro não cadastrada
                If lErro = 5871 Then Error 5875
                
            End If
            
        End If
                
        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
                
        ContaDescricao.Caption = objPlanoConta.sDescConta
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 5697

    Saida_Celula_Conta = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Conta:

    Saida_Celula_Conta = Err
    
    Select Case Err
    
        Case 5696, 5697, 5873, 5874, 20603
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 5806
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", Conta.Text)

            If vbMsgRes = vbYes Then
            
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("PlanoConta", objPlanoConta)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If
            
        Case 5875
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTACCL_INEXISTENTE", Conta.Text, sCcl)

            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
            
                Call Chama_Tela("ContaCcl", objContaCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If

        Case 20639
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162259)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long
'faz a critica da celula ccl do grid que está deixando de ser a corrente

Dim sCclFormatada As String
Dim sContaFormatada As String
Dim lErro As Long
Dim iContaPreenchida As Integer
Dim objContaCcl As New ClassContaCcl
Dim sConta As String
Dim vbMsgRes As VbMsgBoxResult
Dim objCcl As New ClassCcl

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = Ccl
                
    'critica o formato do ccl, sua presença no BD e capacidade de receber lançamentos
    lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
    If lErro <> SUCESSO And lErro <> 5703 Then Error 5707
                
    'se o centro de custo/lucro não estiver cadastrado
    If lErro = 5703 Then Error 5866
                
    'se o centro de custo foi preenchido
    If Len(Ccl.ClipText) > 0 Then
    
        'se a conta foi informada
        If Len(GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Conta_Col)) > 0 Then
    
            'verificar se a associação da conta com o centro de custo em questão está cadastrada
            sConta = GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Conta_Col)
        
            lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
            If lErro <> SUCESSO Then Error 5876
        
            objContaCcl.sConta = sContaFormatada
            objContaCcl.sCcl = sCclFormatada
        
            lErro = CF("ContaCcl_Le", objContaCcl)
            If lErro <> SUCESSO And lErro <> 5871 Then Error 5877
        
            'associação Conta x Centro de Custo/Lucro não cadastrada
            If lErro = 5871 Then Error 5878
        
        End If
                        
        CclDescricao.Caption = objCcl.sDescCcl
             
'        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.ilinhasExistentes Then
'            objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
'        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 5708

    Saida_Celula_Ccl = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = Err
    
    Select Case Err
    
        Case 5707, 5708, 5876, 5877
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 5866
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)

            If vbMsgRes = vbYes Then
            
                objCcl.sCcl = sCclFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("CclTela", objCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If
            
        Case 5878
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTACCL_INEXISTENTE", sConta, Ccl.Text)

            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
            
                Call Chama_Tela("ContaCcl", objContaCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162260)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Credito(objGridInt As AdmGrid) As Long
'faz a critica da celula credito do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dColunaSoma As Double

On Error GoTo Erro_Saida_Celula_Credito

    Set objGridInt.objControle = Credito
    
    If Len(Credito.Text) > 0 Then
    
        lErro = Valor_Critica(Credito.Text)
        If lErro <> SUCESSO Then Error 5710
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 5711
                
    If Len(Credito.Text) > 0 Then
        GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Debito_Col) = ""
              
        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If
        
    dColunaSoma = GridColuna_Soma(iGrid_Debito_Col)
    TotalDebito = Format(dColunaSoma, "Standard")
    dColunaSoma = GridColuna_Soma(iGrid_Credito_Col)
    TotalCredito = Format(dColunaSoma, "Standard")
   
    Saida_Celula_Credito = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Credito:

    Saida_Celula_Credito = Err
    
    Select Case Err
    
        Case 5710, 5711
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162261)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Debito(objGridInt As AdmGrid) As Long
'faz a critica da celula debito do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dColunaSoma As Double

On Error GoTo Erro_Saida_Celula_Debito

    Set objGridInt.objControle = Debito
    
    If Len(Debito.Text) > 0 Then
    
        lErro = Valor_Critica(Debito.Text)
        If lErro <> SUCESSO Then Error 5712
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 5713
                
    If Len(Debito.Text) > 0 Then
    
        GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Credito_Col) = ""
        
        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If
    
    dColunaSoma = GridColuna_Soma(iGrid_Credito_Col)
    TotalCredito = Format(dColunaSoma, "Standard")
    dColunaSoma = GridColuna_Soma(iGrid_Debito_Col)
    TotalDebito = Format(dColunaSoma, "Standard")

    Saida_Celula_Debito = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Debito:

    Saida_Celula_Debito = Err
    
    Select Case Err
    
        Case 5712, 5713
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162262)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_SeqContraPartida(objGridInt As AdmGrid) As Long
'faz a critica da celula sequencial de contra partida do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_SeqContraPartida

    Set objGridInt.objControle = SeqContraPartida
    
    If Len(SeqContraPartida.Text) > 0 Then
    
        If GridLancamentos.Row = CInt(SeqContraPartida.Text) Then Error 20619
        
        If CInt(SeqContraPartida.Text) > objGridInt.iLinhasExistentes Then Error 20620
        
        If CInt(SeqContraPartida.Text) <= 0 Then Error 20621
    
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 20622
                
    Saida_Celula_SeqContraPartida = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_SeqContraPartida:

    Saida_Celula_SeqContraPartida = Err
    
    Select Case Err
    
        Case 20619
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTRAPARTIDA_NAO_MESMO_LANCAMENTO", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 20620, 20621
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTRAPARTIDA_LANCAMENTO_INEXISTENTE", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 20622
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162263)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Historico(objGridInt As AdmGrid) As Long
'faz a critica da celula historico do grid que está deixando de ser a corrente

Dim sValor As String
Dim lErro As Long
Dim objHistPadrao As ClassHistPadrao
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Historico

    Set objHistPadrao = New ClassHistPadrao
    
    Set objGridInt.objControle = Historico
                
    If left(Historico.Text, 1) = CARACTER_HISTPADRAO Then
    
        sValor = Trim(Mid(Historico.Text, 2))
        
        lErro = Valor_Inteiro_Critica(sValor)
        If lErro <> SUCESSO Then Error 5895
        
        objHistPadrao.iHistPadrao = CInt(sValor)
                
        lErro = CF("HistPadrao_Le", objHistPadrao)
        If lErro <> SUCESSO And lErro <> 5446 Then Error 5896
        
        If lErro = 5446 Then Error 5897

        Historico.Text = objHistPadrao.sDescHistPadrao
        
    End If
    
    If Len(Historico.Text) > 0 Then
        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 5714

    Saida_Celula_Historico = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Historico:

    Saida_Celula_Historico = Err
    
    Select Case Err
    
        Case 5714, 5896
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 5895
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_HISTPADRAO_INVALIDO", Err, sValor)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 5897
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_HISTPADRAO_INEXISTENTE", objHistPadrao.iHistPadrao)

            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("HistoricoPadrao", objHistPadrao)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162264)
        
    End Select

    Exit Function

End Function

Function GridColuna_Soma(iColuna As Integer) As Double
    
Dim dAcumulador As Double
Dim iLinha As Integer
    
    dAcumulador = 0
    
    For iLinha = 1 To objGrid1.iLinhasExistentes
        If Len(GridLancamentos.TextMatrix(iLinha, iColuna)) > 0 Then
            dAcumulador = dAcumulador + CDbl(GridLancamentos.TextMatrix(iLinha, iColuna))
        End If
    Next
    
    GridColuna_Soma = dAcumulador

End Function

Private Sub BotaoGravar_Click()

    Call Gravar_Registro
    
    iAlterado = 0
    
End Sub

Public Function Gravar_Registro() As Long
    
Dim lErro As Long
Dim lDoc As Long
Dim colLancamento_Detalhe As New Collection
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim iIndice1 As Integer
Dim dSoma As Double
Dim iPeriodoDoc As Integer
Dim iExercicioDoc As Integer
    
On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Data, determinação dos exercicio e periodo correspondentes
    If Len(Data.ClipText) = 0 Then gError 6050
    
    'Verifica a existencia de pelo menos um lançamento
    If objGrid1.iLinhasExistentes = 0 Then gError 6016
    
    'Lote
    If Len(Lote.ClipText) = 0 Then gError 6002
    
    'Documento
    If Len(Documento.ClipText) = 0 Then gError 6003
        
    'Origem só pode ser CTB
'    If gobjColOrigem.Origem(Origem.Caption) <> "CTB" Then gError 59500
    
    'Preenche Objeto Lançamento_Cabeçalho
    objLancamento_Cabecalho.iFilialEmpresa = giFilialEmpresa
    objLancamento_Cabecalho.sOrigem = gobjColOrigem.Origem(Origem.Caption)
    objLancamento_Cabecalho.iLote = CInt(Lote.ClipText)
    objLancamento_Cabecalho.lDoc = CLng(Documento.ClipText)
    objLancamento_Cabecalho.dtData = CDate(Data.Text)
    
    'Preenche Objeto Lançamento_Detalhe
    lErro = Grid_Lancamento_Detalhe(colLancamento_Detalhe)
    If lErro <> SUCESSO Then gError 5722
        
    'Testa se soma dos débitos é igual a soma dos créditos
    dSoma = 0
    
    For Each objLancamento_Detalhe In colLancamento_Detalhe
        dSoma = dSoma + objLancamento_Detalhe.dValor
    Next
    
    dSoma = Format(dSoma, "Fixed")
    
    If dSoma <> 0 Then gError 6068
            
            
    'Se a Origem for diferente de CTB, só permitir alterar os dados
    If gobjColOrigem.Origem(Origem.Caption) <> "CTB" And gobjColOrigem.Origem(Origem.Caption) <> "FLH" Then
    
        'permite alterar os dados do lançamento
        lErro = CF("LanPendente_Grava", objLancamento_Cabecalho, colLancamento_Detalhe)
        If lErro <> SUCESSO Then gError 83504
    
    Else
            
        lErro = CF("Lancamento_Grava", objLancamento_Cabecalho, colLancamento_Detalhe)
        If lErro <> SUCESSO Then gError 6104
    
    End If
    
    Call Limpa_Tela_Lancamentos

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 5722, 6104, 83504
        
        Case 59500
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_DIFERENTE_CTB", gErr)
        
        Case 6002
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_LOTE_NAO_PREENCHIDO", gErr)
            Lote.SetFocus
        
        Case 6003
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO", gErr)
            Documento.SetFocus
        
        Case 6016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_LANCAMENTOS_GRAVAR", gErr)
        
        Case 6050
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DOCUMENTO_NAO_PREENCHIDA", gErr)
            Data.SetFocus
            
        Case 6068
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_NAO_BALANCEADO", gErr, objLancamento_Cabecalho.lDoc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162265)
            
    End Select
    
    Exit Function
    
End Function

Function Limpa_Tela_Lancamentos() As Long

Dim lErro As Long
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Call Grid_Limpa(objGrid1)
    TotalDebito.Caption = ""
    TotalCredito.Caption = ""
    ContaDescricao.Caption = ""
    CclDescricao.Caption = ""
    DocAuto.Text = ""
    Documento.Text = ""
    Origem.Caption = "Contabilidade"
        
    objGrid1.iProibidoExcluir = 0
    objGrid1.iProibidoIncluir = 0
        
    Limpa_Tela_Lancamentos = SUCESSO
    
End Function

Function Grid_Lancamento_Detalhe(colLancamento_Detalhe As Collection) As Long

Dim iIndice1 As Integer
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim sConta As String
Dim sContaFormatada As String
Dim sCcl As String
Dim sCclFormatada As String
Dim dValorDebito As Double
Dim dValorCredito As Double
Dim iContaPreenchida As Integer
Dim iCclPreenchida As Integer
Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim colContraPartida As New Collection

On Error GoTo Erro_Grid_Lancamento_Detalhe

    For iIndice1 = 1 To objGrid1.iLinhasExistentes
        
        Set objLancamento_Detalhe = New ClassLancamento_Detalhe
        
        objLancamento_Detalhe.iSeq = iIndice1
            
        sConta = GridLancamentos.TextMatrix(iIndice1, iGrid_Conta_Col)
            
        If Len(Trim(sConta)) = 0 Then Error 11245
            
        lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then Error 5701
            
        If iContaPreenchida = CONTA_VAZIA Then Error 9285
            
        objLancamento_Detalhe.sConta = sContaFormatada
    
        'Testa para ver se houve crédito ou débito
        If Len(GridLancamentos.TextMatrix(iIndice1, iGrid_Credito_Col)) > 0 Then
            dValorCredito = CDbl(GridLancamentos.TextMatrix(iIndice1, iGrid_Credito_Col))
        Else
            dValorCredito = 0
        End If
            
        If Len(GridLancamentos.TextMatrix(iIndice1, iGrid_Debito_Col)) > 0 Then
            dValorDebito = CDbl(GridLancamentos.TextMatrix(iIndice1, iGrid_Debito_Col))
        Else
            dValorDebito = 0
        End If
    
        'Armazena débito ou crédito
        If dValorDebito = 0 And dValorCredito = 0 Then Error 6007
            
        objLancamento_Detalhe.dValor = dValorCredito - dValorDebito
    
        'armazena o sequencial de contra partida, se estiver preenchido
        If Len(GridLancamentos.TextMatrix(iIndice1, iGrid_SeqContraPartida_Col)) > 0 Then
            objLancamento_Detalhe.iSeqContraPartida = CInt(GridLancamentos.TextMatrix(iIndice1, iGrid_SeqContraPartida_Col))
            
            lErro = Armazena_Contra_Partida(colContraPartida, objLancamento_Detalhe)
            If lErro <> SUCESSO Then Error 20623
            
        End If
    
        objLancamento_Detalhe.sProduto = ""
    
        'Armazena Histórico e Ccl
        objLancamento_Detalhe.sHistorico = GridLancamentos.TextMatrix(iIndice1, iGrid_Historico_Col)
            
        'verifica se o historico tem parametros que deveriam ter sido substituidos
        If InStr(objLancamento_Detalhe.sHistorico, CARACTER_HISTORICO_PARAM) <> 0 Then Error 20638
            
        'Se está usando Centro de Custo/Lucro, armazena-o
        If iGrid_Ccl_Col <> 999 Then
                
            sCcl = GridLancamentos.TextMatrix(iIndice1, iGrid_Ccl_Col)
            
            lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then Error 5721
            
            If iCclPreenchida = CCL_PREENCHIDA Then
                objLancamento_Detalhe.sCcl = sCclFormatada
            Else
                objLancamento_Detalhe.sCcl = ""
            End If
                
        End If
                
        objLancamento_Detalhe.iGerencial = GridLancamentos.TextMatrix(iIndice1, iGrid_Gerencial_Col)
                
        'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
        colLancamento_Detalhe.Add objLancamento_Detalhe
                
    Next
    
    lErro = Testa_Contra_Partida(colLancamento_Detalhe, colContraPartida)
    If lErro <> SUCESSO Then Error 20624
    
    Grid_Lancamento_Detalhe = SUCESSO

    Exit Function

Erro_Grid_Lancamento_Detalhe:

    Grid_Lancamento_Detalhe = Err

    Select Case Err
    
        Case 5701, 5721, 20623, 20624
        
        Case 6007
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_LANCAMENTO_NAO_PREENCHIDO", Err)
            GridLancamentos.Row = iIndice1
            GridLancamentos.Col = iGrid_Debito_Col
            GridLancamentos.SetFocus
            
        Case 9285, 11245
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_GRID_NAO_PREENCHIDA", Err, iIndice1)
    
        Case 20638
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTORICO_PARAM", Err)
            GridLancamentos.Row = iIndice1
            GridLancamentos.Col = iGrid_Historico_Col
            GridLancamentos.SetFocus
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162266)
            
    End Select
    
    Exit Function

End Function

Private Sub BotaoExcluir_Click()
    
Dim lErro As Long
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim vbMsgRes As VbMsgBoxResult
Dim lDoc As Long

On Error GoTo Erro_BotaoExcluir_Click
     
    GL_objMDIForm.MousePointer = vbHourglass
     
    'Data, determinação dos exercicio e periodo correspondentes
    If Len(Data.ClipText) = 0 Then gError 6105
    
    'Lote
    If Len(Lote.ClipText) = 0 Then gError 6110
    
    'Documento
    If Len(Documento.ClipText) = 0 Then gError 6114
     
    'Origem só pode ser CTB
'    If gobjColOrigem.Origem(Origem.Caption) <> "CTB" Then gError 59511
 
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_DOCUMENTO")
    
    If vbMsgRes = vbYes Then
    
        objLancamento_Cabecalho.iFilialEmpresa = giFilialEmpresa
        objLancamento_Cabecalho.sOrigem = gobjColOrigem.Origem(Origem.Caption)
        objLancamento_Cabecalho.iLote = CInt(Lote.ClipText)
        objLancamento_Cabecalho.lDoc = CLng(Documento.ClipText)
        objLancamento_Cabecalho.dtData = CDate(Data.Text)
        
        lErro = CF("Lancamento_Exclui", objLancamento_Cabecalho)
        If lErro <> SUCESSO Then gError 6120
        
        Call Limpa_Tela_Lancamentos
    
        iAlterado = 0
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
            
        Case 59511
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_DIFERENTE_CTB", gErr)

        Case 6105
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DOCUMENTO_NAO_PREENCHIDA", gErr)
            Data.SetFocus
        
        Case 6110
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_LOTE_NAO_PREENCHIDO", gErr)
            Lote.SetFocus
        
        Case 6114
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO", gErr)
            Documento.SetFocus
        
        Case 6120
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162267)
        
    End Select

    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()

Dim lDoc As Long
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 5931

    Call Limpa_Tela_Lancamentos
               
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 5931
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162268)
        
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub Label2_Click()

Dim objLote As New ClassLote
Dim dtData As Date
Dim lErro As Long
Dim objPeriodo As New ClassPeriodo
Dim colSelecao As New Collection

On Error GoTo Erro_Label2_Click

    'Obtém Periodo e Exercicio correspondentes à data
    If Len(Data.ClipText) > 0 Then
        dtData = CDate(Data.Text)
    
        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 9175
    
    Else
        objPeriodo.iExercicio = 0
        objPeriodo.iPeriodo = 0
    End If
    
    If Len(Lote.Text) = 0 Then
        objLote.iLote = 0
    Else
        objLote.iLote = CInt(Lote.Text)
    End If
    
    objLote.sOrigem = gobjColOrigem.Origem(Origem.Caption)
    objLote.iExercicio = objPeriodo.iExercicio
    objLote.iPeriodo = objPeriodo.iPeriodo
      
    colSelecao.Add giFilialEmpresa
    colSelecao.Add 0
      
    Call Chama_Tela("LotePendenteCTBLista", colSelecao, objLote, objEventoLote)

    Exit Sub

Erro_Label2_Click:
    
    Select Case Err
    
        Case 9175
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162269)
        
    End Select
    
    Exit Sub

End Sub

Private Sub Label3_Click()

Dim objLancamento_Detalhe As New ClassLancamento_Detalhe
Dim dtData As Date
Dim lErro As Long
Dim objPeriodo As New ClassPeriodo
Dim colSelecao As New Collection

On Error GoTo Erro_Label3_Click

    If Len(Data.ClipText) > 0 Then

        'Obtém Periodo e Exercicio correspondentes à data
        dtData = CDate(Data.Text)
    
        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 9257
        
    Else
        objPeriodo.iExercicio = 0
        objPeriodo.iPeriodo = 0
    End If
    
    If Len(Documento.Text) = 0 Then
        objLancamento_Detalhe.lDoc = 0
    Else
        objLancamento_Detalhe.lDoc = CLng(Documento.ClipText)
    End If
    
    objLancamento_Detalhe.iFilialEmpresa = giFilialEmpresa
    objLancamento_Detalhe.sOrigem = gobjColOrigem.Origem(Origem.Caption)
    objLancamento_Detalhe.iExercicio = objPeriodo.iExercicio
    objLancamento_Detalhe.iPeriodoLan = objPeriodo.iPeriodo
    objLancamento_Detalhe.iPeriodoLote = objPeriodo.iPeriodo
    
    If Len(Lote.Text) = 0 Then
        objLancamento_Detalhe.iLote = 0
    Else
        objLancamento_Detalhe.iLote = CInt(Lote.Text)
    End If
    
    objLancamento_Detalhe.iSeq = 0
    
    Call Chama_Tela("LanPendenteLista", colSelecao, objLancamento_Detalhe, objEventoLancamento)

    Exit Sub

Erro_Label3_Click:
    
    Select Case Err
    
        Case 9257
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162270)
        
    End Select
    
    Exit Sub

End Sub

Private Sub Label6_Click()

Dim objDocAuto As New ClassDocAuto
Dim colSelecao As Collection

    If Len(DocAuto.Text) = 0 Then
        objDocAuto.lDoc = 0
    Else
        objDocAuto.lDoc = CLng(DocAuto.ClipText)
    End If

    objDocAuto.iSeq = 0

    Call Chama_Tela("DocAutoLista", colSelecao, objDocAuto, objEventoDocAuto)

End Sub

Private Sub Lote_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Data_Validate(Cancel As Boolean)
    
Dim lErro As Long
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim lDoc As Long
Dim sNomeExterno As String
Dim objLote As New ClassLote
Dim vbMsgRes As VbMsgBoxResult
Dim iLoteAtualizado As Integer
Dim objPeriodosFilial As New ClassPeriodosFilial

On Error GoTo Erro_Data_Validate

    If Len(Data.ClipText) > 0 Then

        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then Error 6092
    
        'Obtém Periodo e Exercicio correspondentes à data
        dtData = CDate(Data.Text)
    
        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 6045
    
        'Verifica se Exercicio está fechado
        lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
        If lErro <> SUCESSO And lErro <> 10083 Then Error 6096
        
        'Exercicio não cadastrado
        If lErro = 10083 Then Error 10086
        
        If objExercicio.iStatus = EXERCICIO_FECHADO Then Error 6094
        
        objPeriodosFilial.iFilialEmpresa = giFilialEmpresa
        objPeriodosFilial.iExercicio = objPeriodo.iExercicio
        objPeriodosFilial.iPeriodo = objPeriodo.iPeriodo
        objPeriodosFilial.sOrigem = gobjColOrigem.Origem(Origem.Caption)
        
        lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
        If lErro <> SUCESSO Then Error 10160
        
        If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 10161
        
        'checa se o lote pertence ao periodo em questão
        If Len(Lote.Text) > 0 Then
    
            objLote.iLote = CInt(Lote.Text)
    
            objLote.iFilialEmpresa = giFilialEmpresa
            objLote.sOrigem = gobjColOrigem.Origem(Origem.Caption)
            objLote.iExercicio = objPeriodo.iExercicio
            objLote.iPeriodo = objPeriodo.iPeriodo
    
            'verifica se o lote  está atualizado
            lErro = CF("Lote_Critica_Atualizado", objLote, iLoteAtualizado)
            If lErro <> SUCESSO Then Error 5997
    
            'Se é um lote que já foi contabilizado, não pode sofrer alteração
            If iLoteAtualizado = LOTE_ATUALIZADO Then Error 5906
    
            lErro = CF("LotePendente_Le", objLote)
            If lErro <> SUCESSO And lErro <> 5435 Then Error 5903
    
            'Se o lote não está cadastrado
            If lErro = 5435 Then Error 5904
        
            If giSetupLotePorPeriodo <> LOTE_INICIALIZADO_POR_PERIODO And objPeriodo.iPeriodo <> objLote.iPeriodo Then Error 5905
    
        End If
        
        'Preenche campo de periodo
        Periodo.Caption = objPeriodo.sNomeExterno
    
        Exercicio.Caption = objExercicio.sNomeExterno
    
    Else
    
        Periodo.Caption = ""
    
        Exercicio.Caption = ""
        
    End If
        
    Exit Sub

Erro_Data_Validate:
    
    Cancel = True
    
    If Not (Parent Is GL_objMDIForm.ActiveForm) Then
        Me.Show
    End If
    
    Select Case Err
    
        Case 5903, 5997, 5900
    
        Case 5904
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_INEXISTENTE", Err, objLote.iFilialEmpresa, objLote.iLote, objLote.iExercicio, objLote.iPeriodo, Origem.Caption)
            
            If vbMsgRes = vbYes Then
                'Se respondeu que deseja criar LOTE
                Call Chama_Tela("LoteTela", objLote)
            End If
    
        Case 5905
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODOS_DIFERENTES", Err, objPeriodo.iPeriodo, objLote.iPeriodo)
    
        Case 5906
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_ATUALIZADO_NAO_RECEBE_LANCAMENTOS", Err, objLote.iFilialEmpresa, objLote.iLote, objLote.iExercicio, objLote.iPeriodo, Origem.Caption)
    
        Case 6045, 6096, 10160
    
        Case 6092
           
        Case 6094
            'Não é possível fazer lançamentos em exercício fechado
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_EXERCICIO_FECHADO", Err, objLote.iExercicio)
            
        Case 10086
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)
            
        Case 10161
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", Err, objPeriodosFilial.iExercicio, objPeriodosFilial.iPeriodo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162271)
        
    End Select
    
    Exit Sub

End Sub

Private Sub Lote_Validate(Cancel As Boolean)
    
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim objLote As New ClassLote
Dim sNomeExterno As String
Dim objExercicio As New ClassExercicio
Dim iLoteAtualizado As Integer
Dim colSelecao As Collection
    
On Error GoTo Erro_Lote_Validate

    If Len(Lote.Text) > 0 And Len(Data.ClipText) > 0 Then
    
        objLote.iLote = CInt(Lote.Text)
        objLote.iFilialEmpresa = giFilialEmpresa
        objLote.sOrigem = gobjColOrigem.Origem(Origem.Caption)

        'Obtém Periodo e Exercicio correspondentes à data
        dtData = CDate(Data.Text)
    
        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 5818
        
        objLote.iExercicio = objPeriodo.iExercicio
        objLote.iPeriodo = objPeriodo.iPeriodo
    
        'verifica se o lote  está atualizado
        lErro = CF("Lote_Critica_Atualizado", objLote, iLoteAtualizado)
        If lErro <> SUCESSO Then Error 5998
    
        'Se é um lote que já foi contabilizado, não pode sofrer alteração
        If iLoteAtualizado = LOTE_ATUALIZADO Then Error 6064
    
        lErro = CF("LotePendente_Le", objLote)
        If lErro <> SUCESSO And lErro <> 5435 Then Error 5902
    
        'Se o lote não está cadastrado
        If lErro = 5435 Then Error 6042
        
        If giSetupLotePorPeriodo <> LOTE_INICIALIZADO_POR_PERIODO And objPeriodo.iPeriodo <> objLote.iPeriodo Then Error 5828
    
        
    End If
    
    Exit Sub

Erro_Lote_Validate:

    Cancel = True

    If Not (Parent Is GL_objMDIForm.ActiveForm) Then
        Me.Show
    End If

    Select Case Err
    
        Case 5818, 5902, 5998
                
        Case 5828
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODOS_DIFERENTES", Err, objPeriodo.iPeriodo, objLote.iPeriodo)
    
        Case 6042
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_INEXISTENTE", objLote.iFilialEmpresa, objLote.iLote, Origem.Caption, objPeriodo.iPeriodo, objPeriodo.iExercicio)
            
            If vbMsgRes = vbYes Then
                'Se respondeu que deseja criar LOTE
                Call Chama_Tela("LoteTela", objLote)
            End If
            
        Case 6064
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_ATUALIZADO_NAO_RECEBE_LANCAMENTOS", Err, objLote.iFilialEmpresa, objLote.iLote, objPeriodo.iExercicio, objPeriodo.iPeriodo, Origem.Caption)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162272)
            
    End Select
    
    Exit Sub

End Sub

Public Sub menulimpar_Click()

    Call Grid_Limpa(objGrid1)
    TotalDebito.Caption = ""
    TotalCredito.Caption = ""
    ContaDescricao.Caption = ""
    CclDescricao.Caption = ""

End Sub

Public Sub menuRateio_Click()
'Quando a opção Rateio e Selecionada o menu ele chama a tela Rateio

Dim objRateioOn As New ClassRateioOn
Dim lErro As Long
Dim objConfirmaTela As New AdmConfirmaTela

On Error GoTo Erro_menuRateio_Click

    Call Chama_Tela_Modal("Rateio", objRateioOn, objConfirmaTela)
    
    If objConfirmaTela.iTelaOK = OK Then
    
        lErro = Traz_Rateio_Tela(objRateioOn)
        If lErro <> SUCESSO Then Error 9639
        
    End If
    
    Exit Sub

Erro_menuRateio_Click:

    Select Case Err
    
        Case 9639
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162273)
            
    End Select

    Exit Sub

End Sub

Private Sub objEventoDocAuto_evSelecao(obj1 As Object)
'traz o documento automatico selecionado para a tela
    
    Dim objDocAuto As ClassDocAuto
    Dim lErro As Long
    
On Error GoTo Erro_objEventoDocAuto_evSelecao
    
    Set objDocAuto = obj1
    
    lErro = Traz_DocAuto_Tela(objDocAuto.lDoc)
    If lErro <> SUCESSO Then Error 11120
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoDocAuto_evSelecao:

    Select Case Err
    
        Case 11120
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162274)
            
    End Select
        
    Exit Sub
        
End Sub

Private Sub objEventoLancamento_evSelecao(obj1 As Object)
'traz o lançamento selecionado para a tela

Dim lErro As Long
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim iIndice As Integer
Dim sDescricao As String
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho

On Error GoTo Erro_objEventoLancamento_evSelecao

    Set objLancamento_Detalhe = obj1
    
    objLancamento_Cabecalho.iFilialEmpresa = objLancamento_Detalhe.iFilialEmpresa
    objLancamento_Cabecalho.sOrigem = objLancamento_Detalhe.sOrigem
    objLancamento_Cabecalho.iExercicio = objLancamento_Detalhe.iExercicio
    objLancamento_Cabecalho.iPeriodoLan = objLancamento_Detalhe.iPeriodoLan
    objLancamento_Cabecalho.lDoc = objLancamento_Detalhe.lDoc
    
    lErro = Traz_Doc_Tela(objLancamento_Cabecalho)
    If lErro <> SUCESSO And lErro <> 5843 Then Error 9258
    
    'documento não cadastrado
    If lErro = 5843 Then Error 20296
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
        
    Me.Show
    
    Exit Sub

Erro_objEventoLancamento_evSelecao:
    
    Select Case Err
    
        Case 9258
    
        Case 20296
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOC_NAO_CADASTRADO", Err, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162275)
        
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)
'traz o lote selecionado para a tela

Dim lErro As Long
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim objLote As ClassLote
Dim iIndice As Integer
Dim sDescricao As String

On Error GoTo Erro_objEventoLote_evSelecao

    Set objLote = obj1
    
    'Se estiver com a data preenchida ==> verificar se a data está dentro do periodo do lote
    If Len(Data.ClipText) > 0 Then

        'Obtém Periodo e Exercicio correspondentes à data
        dtData = CDate(Data.Text)
    
        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 9189
    
        'se o periodo/exercicio não corresponde ao periodo/exercicio do lote ==> troca a data
        If objPeriodo.iExercicio <> objLote.iExercicio Or objPeriodo.iPeriodo <> objLote.iPeriodo Then
                        
            'move a data inicial do lote, exercicio e periodo para a tela
            lErro = Move_Data_Tela(objLote)
            If lErro <> SUCESSO Then Error 9187
        
        End If
        
    Else
    
        'se não estiver com a data preenchida
        'move a data inicial do lote, exercicio e periodo para a tela
        lErro = Move_Data_Tela(objLote)
        If lErro <> SUCESSO Then Error 9188
               
    End If
    
    Lote.Text = CStr(objLote.iLote)
    Origem.Caption = gobjColOrigem.Descricao(objLote.sOrigem)
        
    Me.Show
    
    Exit Sub

Erro_objEventoLote_evSelecao:
    
    Select Case Err
    
        Case 9187, 9188, 9189  'Erro já tratado na rotina chamada
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162276)
        
    End Select
    
    Exit Sub
    
End Sub

Private Function Move_Data_Tela(objLote As ClassLote) As Long

Dim lErro As Long
Dim objExercicio As New ClassExercicio
Dim objPeriodo As New ClassPeriodo
Dim objPeriodosFilial As New ClassPeriodosFilial

On Error GoTo Erro_Move_Data_Tela
    
    lErro = CF("Periodo_Le_ExercicioPeriodo", objLote.iExercicio, objLote.iPeriodo, objPeriodo)
    If lErro <> SUCESSO Then Error 9190

    'Verifica se Exercicio está fechado
    lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
    If lErro <> SUCESSO And lErro <> 10083 Then Error 9191
    
    'se o exercicio não estiver cadastrado
    If lErro = 10083 Then Error 10087
                        
    If objExercicio.iStatus = EXERCICIO_FECHADO Then Error 10162
    
    objPeriodosFilial.iFilialEmpresa = giFilialEmpresa
    objPeriodosFilial.iExercicio = objPeriodo.iExercicio
    objPeriodosFilial.iPeriodo = objPeriodo.iPeriodo
    objPeriodosFilial.sOrigem = gobjColOrigem.Origem(Origem.Caption)
    
    lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
    If lErro <> SUCESSO Then Error 10163
    
    If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 10164
                        
    Data.Text = Format(objPeriodo.dtDataInicio, "dd/mm/yy")

    Periodo.Caption = objPeriodo.sNomeExterno
    
    Exercicio.Caption = objExercicio.sNomeExterno
    
    Move_Data_Tela = SUCESSO
    
    Exit Function

Erro_Move_Data_Tela:
    
    Move_Data_Tela = Err
    
    Select Case Err
    
        Case 9190, 9191, 10163
    
        Case 10087
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)
            
        Case 10162
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_EXERCICIO_FECHADO", Err, objPeriodo.iExercicio)
            
        Case 10164
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", Err, objPeriodosFilial.iExercicio, objPeriodosFilial.iPeriodo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162277)
        
    End Select
    
    Exit Function

End Function

Private Sub TvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
    
Dim sCcl As String
Dim sCclEnxuta As String
Dim lErro As Long
Dim lPosicaoSeparador As Long
Dim sCaracterInicial As String
    
On Error GoTo Erro_TvwCcls_NodeClick
    
    If GridLancamentos.Col = iGrid_Ccl_Col Then
    
        sCaracterInicial = left(Node.Key, 1)
    
        If sCaracterInicial = "A" Then
    
            sCcl = right(Node.Key, Len(Node.Key) - 1)
              
            sCclEnxuta = String(STRING_CCL, 0)
            
            'volta mascarado apenas os caracteres preenchidos
            lErro = Mascara_RetornaCclEnxuta(sCcl, sCclEnxuta)
            If lErro <> SUCESSO Then Error 10499
            
            Ccl.PromptInclude = False
            Ccl.Text = sCclEnxuta
            Ccl.PromptInclude = True
              
            GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col) = Ccl.Text
        
            If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.iLinhasExistentes Then
                objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
            End If
        
            'Preenche a Descricao do centro de custo/lucro
            lPosicaoSeparador = InStr(Node.Text, SEPARADOR)
            CclDescricao.Caption = Mid(Node.Text, lPosicaoSeparador + 1)
    
        End If
    
    End If
    
    Exit Sub

Erro_TvwCcls_NodeClick:

    Select Case Err
    
        Case 10499
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", Err, sCcl)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162278)
            
    End Select
        
    Exit Sub

End Sub

Private Sub ListDocAuto_DblClick()

Dim lDoc As Long
Dim lErro As Long
Dim objDocAuto As New ClassDocAuto
Dim lPosicaoSeparador As Long

On Error GoTo Erro_ListDocAuto_DlbClick
    
    'Guarda a posicao em que o separador se encontra
    lPosicaoSeparador = InStr(ListDocAuto.Text, SEPARADOR)
    
    'Pega o Numero do Documento Automático selecionado
    lDoc = CLng(left(ListDocAuto.Text, lPosicaoSeparador - 1))
    
    objDocAuto.lDoc = lDoc

    lErro = Traz_DocAuto_Tela(lDoc)
    If lErro <> SUCESSO Then Error 11351
    
    Exit Sub
    
Erro_ListDocAuto_DlbClick:

    Select Case Err
    
        Case 11351
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 162279)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub ListHistoricos_DblClick()

Dim lPosicaoSeparador As Long
    
    If GridLancamentos.Col = iGrid_Historico_Col Then
    
        lPosicaoSeparador = InStr(ListHistoricos.Text, SEPARADOR)
        GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col) = Mid(ListHistoricos.Text, lPosicaoSeparador + 1)
        Historico.Text = Mid(ListHistoricos.Text, lPosicaoSeparador + 1)
    
        If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.iLinhasExistentes Then
            objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
        End If
    
    End If
    
End Sub

Private Sub TvwContas_Click()

    TvwContas.Tag = "1"
    
End Sub

Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
    
Dim lErro As Long
Dim iIndice As Integer
Dim sConta As String
Dim sCaracterInicial As String
Dim lPosicaoSeparador As Long
Dim sContaEnxuta As String
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_TvwContas_NodeClick
    
    If GridLancamentos.Col = iGrid_Conta_Col Then
    
        sCaracterInicial = left(Node.Key, 1)
    
        If sCaracterInicial = "A" Then
    
            sConta = right(Node.Key, Len(Node.Key) - 1)
            
            sContaEnxuta = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
            If lErro <> SUCESSO Then Error 5855
            
            Conta.PromptInclude = False
            Conta.Text = sContaEnxuta
            Conta.PromptInclude = True
        
            GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col) = Conta.Text
        
            If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.iLinhasExistentes Then
                objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
            End If
        
            'Preenche a Descricao da Conta
            lPosicaoSeparador = InStr(Node.Text, SEPARADOR)
            ContaDescricao.Caption = Mid(Node.Text, lPosicaoSeparador + 1)
            
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", Conta.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
            If lErro <> SUCESSO And lErro <> 5700 Then Error 19137
                    
            'Conta não cadastrada
            If lErro = 5700 Then Error 19138
        
            'Se a Conta possui um Histórico Padrão "default" coloca na tela
            If Len(Trim(GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Historico_Col))) = 0 And objPlanoConta.iHistPadrao <> 0 Then
                            
                For iIndice = 0 To ListHistoricos.ListCount - 1
                    If ListHistoricos.ItemData(iIndice) = objPlanoConta.iHistPadrao Then
                        ListHistoricos.ListIndex = iIndice
                        lPosicaoSeparador = InStr(ListHistoricos.Text, SEPARADOR)
                        GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Historico_Col) = Mid(ListHistoricos.Text, lPosicaoSeparador + 1)
                        Exit For
                    End If
                Next
        
            End If
        
        End If
        
    End If
        
    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case Err
    
        Case 5855
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)
    
        Case 19137
    
        Case 11938
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, Conta.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162280)
            
    End Select
        
    Exit Sub
    
End Sub

Function Conta_Exibe_Descricao(sConta As String) As Long
'exibe a descrição da conta no campo ContaDescricao. A conta passada como parametro deve estar mascarada

Dim sContaFormatada As String
Dim lErro As Long
Dim iContaPreenchida As Integer
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_Conta_Exibe_Descricao

    'Retorna conta formatada como no BD
    lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
    If lErro <> SUCESSO Then Error 5804
    
    lErro = CF("Conta_SelecionaUma", sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
    If lErro <> SUCESSO And lErro <> 6030 Then Error 5881

    If lErro = 6030 Then Error 5883
    
    ContaDescricao.Caption = objPlanoConta.sDescConta
    
    Conta_Exibe_Descricao = SUCESSO
    
    Exit Function

Erro_Conta_Exibe_Descricao:

    Conta_Exibe_Descricao = Err
    
    Select Case Err
    
        Case 5804, 5881
            ContaDescricao = ""
            
        Case 5883
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, sConta)
            ContaDescricao = ""
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162281)
            
    End Select
        
    Exit Function

End Function

Function Ccl_Exibe_Descricao(sCcl As String) As Long
'exibe a descrição do centro de custo/lucro no campo CclDescricao. O ccl passado como parametro deve estar mascarado

Dim sCclFormatada As String
Dim sCclArvore As String
Dim objNode As Node
Dim lErro As Long
Dim iCclPreenchida As Integer
Dim objCcl As New ClassCcl

On Error GoTo Erro_Ccl_Exibe_Descricao

    'Retorna Ccl formatada como no BD
    lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
    If lErro <> SUCESSO Then Error 5805
    
    objCcl.sCcl = sCclFormatada

    lErro = CF("Ccl_Le", objCcl)
    If lErro <> SUCESSO And lErro <> 5599 Then Error 5884
    
    If lErro = 5599 Then Error 5885
    
    CclDescricao.Caption = objCcl.sDescCcl
    
    Ccl_Exibe_Descricao = SUCESSO
    
    Exit Function

Erro_Ccl_Exibe_Descricao:

    Ccl_Exibe_Descricao = Err
    
    Select Case Err
    
        Case 5805, 5884
            CclDescricao = ""
            
        Case 5885
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", Err, objCcl.sCcl)
            CclDescricao = ""
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162282)
            
    End Select
        
    Exit Function

End Function

Private Sub UpDown1_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_DownClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text
        
        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 5925
        
        Data.Text = sData
        
    End If
    
    Exit Sub
    
Erro_UpDown1_DownClick:
    
    Select Case Err
    
        Case 5925
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162283)
        
    End Select
    
    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_UpClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 5926
        
        Data.Text = sData
        
    End If
    
    Exit Sub
    
Erro_UpDown1_UpClick:
    
    Select Case Err
    
        Case 5926
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162284)
        
    End Select
    
    Exit Sub

End Sub

Private Function Traz_DocAuto_Tela(lDoc As Long) As Long
'Coloca no GridLancamentos os Lancamentos do Documento Automático lDoc
    
Dim colDocAuto As New Collection
Dim objDocAuto As New ClassDocAuto
Dim objDocAuto1 As ClassDocAuto
Dim dTotalCredito As Double
Dim dTotalDebito As Double
Dim sContaMascarada As String
Dim sCclMascarado As String
Dim sDescricao As String
Dim lErro As Long
Dim sDoc As String
Dim iLinha As Integer

On Error GoTo Erro_Traz_DocAuto_Tela
    
    Set objDocAuto = New ClassDocAuto
    
    objDocAuto.lDoc = lDoc
    
    'Le os dados do Documento Automatico passado como parâmetro
    lErro = CF("DocAuto_Le_Doc", objDocAuto, colDocAuto)
    If lErro <> SUCESSO And lErro <> 11017 Then Error 11349
    
    'se não encontrou o documento
    If lErro = 11017 Then Error 11350
    
    sDoc = Documento.Text
    
    Call Limpa_Tela_Lancamentos

    Documento.Text = sDoc
    
    'Transfere os dados para a Tela
    DocAuto.Text = CStr(lDoc)
    
    For Each objDocAuto1 In colDocAuto

        iLinha = iLinha + 1

        If Len(objDocAuto1.sConta) > 0 Then
        
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objDocAuto1.sConta, sContaMascarada)
            If lErro <> SUCESSO Then Error 11358
            
            Conta.PromptInclude = False
            Conta.Text = sContaMascarada
            Conta.PromptInclude = True
        
            'coloca a conta na tela
            GridLancamentos.TextMatrix(iLinha, iGrid_Conta_Col) = Conta.Text
            
        Else
        
            'coloca a conta na tela
            GridLancamentos.TextMatrix(iLinha, iGrid_Conta_Col) = ""
        
        End If
        
        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        
            If Len(objDocAuto1.sCcl) > 0 Then
        
                'mascara o centro de custo
                sCclMascarado = String(STRING_CCL, 0)
            
                lErro = Mascara_RetornaCclEnxuta(objDocAuto1.sCcl, sCclMascarado)
                If lErro <> SUCESSO Then Error 11359
            
                Ccl.PromptInclude = False
                Ccl.Text = sCclMascarado
                Ccl.PromptInclude = True
            
                'coloca o centro de custo na tela
                GridLancamentos.TextMatrix(iLinha, iGrid_Ccl_Col) = Ccl.Text
         
            Else
            
                GridLancamentos.TextMatrix(iLinha, iGrid_Ccl_Col) = ""
            
            End If
            
        End If
        
        'coloca o valor na tela
        If objDocAuto1.dValor > 0 Then
            GridLancamentos.TextMatrix(iLinha, iGrid_Credito_Col) = Format(objDocAuto1.dValor, "Standard")
            dTotalCredito = dTotalCredito + objDocAuto1.dValor
        Else
            GridLancamentos.TextMatrix(iLinha, iGrid_Debito_Col) = Format(-objDocAuto1.dValor, "Standard")
            dTotalDebito = dTotalDebito - objDocAuto1.dValor
        End If
            
        'coloca o sequencial de contra-partida na tela
        If objDocAuto1.iSeqContraPartida <> 0 Then GridLancamentos.TextMatrix(iLinha, iGrid_SeqContraPartida_Col) = CStr(objDocAuto1.iSeqContraPartida)
            
        'coloca o historico na tela
        GridLancamentos.TextMatrix(iLinha, iGrid_Historico_Col) = objDocAuto1.sHistorico
            
        objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
    
    Next

    TotalCredito.Caption = Format(dTotalCredito, "Standard")
    TotalDebito.Caption = Format(dTotalDebito, "Standard")
    
    Traz_DocAuto_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_DocAuto_Tela:

    Traz_DocAuto_Tela = Err
        
    Select Case Err
    
        Case 11349
        
        Case 11350
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOCAUTO_NAO_CADASTRADO", Err)
            
        Case 11358
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objDocAuto1.sConta)
        
        Case 11359
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", Err, objDocAuto1.sCcl)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162285)
        
    End Select
    
    Exit Function
    
End Function

Function Traz_Rateio_Tela(objRateioOn1 As ClassRateioOn) As Long
'Traz os dados de um rateio, previamente escolhido, para a tela
    
Dim ColRateioOn As New Collection
Dim objRateioOn As New ClassRateioOn
Dim dValorFinal As Double
Dim dTotalDebito As Double
Dim dTotalCredito As Double
Dim sContaMascarada As String
Dim sCclMascarado As String
Dim iUltimaLinha As Integer
Dim lErro As Long
   
On Error GoTo Erro_Traz_Rateio_Tela

    objRateioOn.iCodigo = objRateioOn1.iCodigo
    
    lErro = CF("RateioOn_Le_Doc", objRateioOn, ColRateioOn)
    If lErro <> SUCESSO And lErro <> 11136 Then Error 11365
    
    'se não encontrou o Rateio
    If lErro = 11136 Then Error 11366
        
    'Variarel que vai indicar apartir de onde deve ser inserido os lancamentos vindos na collection
    iUltimaLinha = objGrid1.iLinhasExistentes
        
    For Each objRateioOn In ColRateioOn

        If Len(objRateioOn.sConta) > 0 Then

            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
        
            lErro = Mascara_RetornaContaEnxuta(objRateioOn.sConta, sContaMascarada)
            If lErro <> SUCESSO Then Error 20294
        
            Conta.PromptInclude = False
            Conta.Text = sContaMascarada
            Conta.PromptInclude = True
            
            'coloca a conta na tela
            GridLancamentos.TextMatrix(objRateioOn.iSeq + iUltimaLinha, iGrid_Conta_Col) = Conta.Text
                
        Else
            
            GridLancamentos.TextMatrix(objRateioOn.iSeq + iUltimaLinha, iGrid_Conta_Col) = ""
            
        End If
            
        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        
            If Len(objRateioOn.sCcl) > 0 Then
            
                'mascara o centro de custo
                sCclMascarado = String(STRING_CCL, 0)
            
                'mascara o centro de custo
                sCclMascarado = String(STRING_CCL, 0)
            
                lErro = Mascara_RetornaCclEnxuta(objRateioOn.sCcl, sCclMascarado)
                If lErro <> SUCESSO Then Error 20295
            
                Ccl.PromptInclude = False
                Ccl.Text = sCclMascarado
                Ccl.PromptInclude = True
            
                'coloca o centro de custo na tela
                GridLancamentos.TextMatrix(objRateioOn.iSeq + iUltimaLinha, iGrid_Ccl_Col) = Ccl.Text
            
            Else
                
                GridLancamentos.TextMatrix(objRateioOn.iSeq + iUltimaLinha, iGrid_Ccl_Col) = ""
                
            End If
            
        End If
        
        dValorFinal = objRateioOn1.dPercentual * objRateioOn.dPercentual
        
        If dValorFinal > 0 Then
            GridLancamentos.TextMatrix(objRateioOn.iSeq + iUltimaLinha, iGrid_Credito_Col) = Format(dValorFinal, "Standard")
        Else
            GridLancamentos.TextMatrix(objRateioOn.iSeq + iUltimaLinha, iGrid_Debito_Col) = Format(-dValorFinal, "Standard")
        End If
            
        'coloca o historico na tela
        GridLancamentos.TextMatrix(objRateioOn.iSeq + iUltimaLinha, iGrid_Historico_Col) = objRateioOn.sHistorico
            
        objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
            
    Next

    dTotalCredito = GridColuna_Soma(iGrid_Credito_Col)
    dTotalDebito = GridColuna_Soma(iGrid_Debito_Col)
    
    TotalCredito.Caption = Format(dTotalCredito, "Standard")
    TotalDebito.Caption = Format(dTotalDebito, "Standard")
    
    Traz_Rateio_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Rateio_Tela:

    Traz_Rateio_Tela = Err
        
    Select Case Err
    
        Case 11365
        
        Case 11366
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RATEIOON_NAO_CADASTRADO", Err, objRateioOn1.iCodigo)
    
        Case 20294
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objRateioOn.sConta)
        
        Case 20295
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", Err, objRateioOn.sCcl)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162286)
        
    End Select
    
    Exit Function
    
End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim colLancamento_Detalhe As New Collection
Dim objExercicio As New ClassExercicio
Dim objPeriodo As New ClassPeriodo

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Lanc"
        
    'Data, determinação dos exercicio e periodo correspondentes
    If Len(Data.ClipText) = 0 Then
        objLancamento_Cabecalho.dtData = DATA_NULA
    Else
        objLancamento_Cabecalho.dtData = CDate(Data.Text)
    End If
    
    'Lote
    If Len(Trim(Lote.ClipText)) = 0 Then
        objLancamento_Cabecalho.iLote = 0
    Else
        objLancamento_Cabecalho.iLote = CInt(Lote.ClipText)
    End If
    
    'Documento
    If Len(Trim(Documento.ClipText)) = 0 Then
        objLancamento_Cabecalho.lDoc = 0
    Else
        objLancamento_Cabecalho.lDoc = CLng(Documento.ClipText)
    End If
    
    If Len(Trim(Exercicio.Caption)) > 0 Then objExercicio.sNomeExterno = Exercicio.Caption
    
    'Lê o Exercício
    lErro = CF("Exercicio_Le_Codigo", objExercicio)
    If lErro <> SUCESSO And lErro <> 28732 Then Error 28755
    
    If lErro = 28732 Then Error 28756
    
    If objExercicio.iExercicio <> 0 Then objPeriodo.iExercicio = objExercicio.iExercicio
    If Len(Trim(Periodo.Caption)) > 0 Then objPeriodo.sNomeExterno = Periodo.Caption
    
    'Lê o Período
    lErro = CF("Periodo_Le_Codigo", objPeriodo)
    If lErro <> SUCESSO And lErro <> 28736 Then Error 28757
    
    If lErro = 28736 Then Error 28758
    
    objLancamento_Cabecalho.iFilialEmpresa = giFilialEmpresa
    objLancamento_Cabecalho.sOrigem = gobjColOrigem.Origem(Origem.Caption)
    objLancamento_Cabecalho.iExercicio = objExercicio.iExercicio
    objLancamento_Cabecalho.iPeriodoLote = objPeriodo.iPeriodo
    objLancamento_Cabecalho.iPeriodoLan = objPeriodo.iPeriodo
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "FilialEmpresa", objLancamento_Cabecalho.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Origem", objLancamento_Cabecalho.sOrigem, STRING_ORIGEM, "Origem"
    colCampoValor.Add "Data", objLancamento_Cabecalho.dtData, 0, "Data"
    colCampoValor.Add "Lote", objLancamento_Cabecalho.iLote, 0, "Lote"
    colCampoValor.Add "Doc", objLancamento_Cabecalho.lDoc, 0, "Doc"
    colCampoValor.Add "Exercicio", objLancamento_Cabecalho.iExercicio, 0, "Exercicio"
    colCampoValor.Add "PeriodoLan", objLancamento_Cabecalho.iPeriodoLan, 0, "PeriodoLan"
    colCampoValor.Add "PeriodoLote", objLancamento_Cabecalho.iPeriodoLote, 0, "PeriodoLote"
   
    'Exemplo de Filtro para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case Err

        Case 28755, 28757
        
        Case 28756
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_INEXISTENTE", Err, objExercicio.sNomeExterno)
            
        Case 28758
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_EXERCICIO_INEXISTENTE", Err, objPeriodo.iExercicio, objPeriodo.sNomeExterno)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162287)

    End Select
    
    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho

On Error GoTo Erro_Tela_Preenche

    objLancamento_Cabecalho.dtData = colCampoValor.Item("Data").vValor

    If objLancamento_Cabecalho.dtData <> 0 Then
    
        objLancamento_Cabecalho.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
        objLancamento_Cabecalho.sOrigem = colCampoValor.Item("Origem").vValor
        objLancamento_Cabecalho.iLote = colCampoValor.Item("Lote").vValor
        objLancamento_Cabecalho.lDoc = colCampoValor.Item("Doc").vValor
        objLancamento_Cabecalho.iExercicio = colCampoValor.Item("Exercicio").vValor
        objLancamento_Cabecalho.iPeriodoLan = colCampoValor.Item("PeriodoLan").vValor
        objLancamento_Cabecalho.iPeriodoLote = colCampoValor.Item("PeriodoLote").vValor
        
        lErro = Traz_Doc_Tela(objLancamento_Cabecalho)
        If lErro <> SUCESSO And lErro <> 5843 Then Error 14965

        If lErro = 5843 Then Error 20297

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 14965

        Case 20297
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOC_NAO_CADASTRADO", Err, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162288)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_TvwContas_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then
    
        'move os dados do plano de contas do banco de dados para a arvore colNodes.
        lErro = CF("Carga_Arvore_Conta1", objNode, TvwContas.Nodes)
        If lErro <> SUCESSO Then Error 44023
        
    End If
    
    Exit Sub
    
Erro_TvwContas_Expand:

    Select Case Err
    
        Case 44023
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162289)
        
    End Select
        
    Exit Sub
    
End Sub

Private Sub Trata_SeqContraPartida(iLinhaExcluida As Integer)
'altera os indicadores de contra partida quando uma linha é excluida

Dim iLinha As Integer

    For iLinha = 1 To objGrid1.iLinhasExistentes
    
        If Len(GridLancamentos.TextMatrix(iLinha, iGrid_SeqContraPartida_Col)) > 0 Then
    
            If CInt(GridLancamentos.TextMatrix(iLinha, iGrid_SeqContraPartida_Col)) = iLinhaExcluida Then
                GridLancamentos.TextMatrix(iLinha, iGrid_SeqContraPartida_Col) = ""
            ElseIf CInt(GridLancamentos.TextMatrix(iLinha, iGrid_SeqContraPartida_Col)) > iLinhaExcluida Then
                GridLancamentos.TextMatrix(iLinha, iGrid_SeqContraPartida_Col) = CStr(CInt(GridLancamentos.TextMatrix(iLinha, iGrid_SeqContraPartida_Col)) - 1)
                
            End If
        End If
    
    Next

End Sub

Private Function Armazena_Contra_Partida(colContraPartida As Collection, objLancamento_Detalhe As ClassLancamento_Detalhe) As Long
'armazena os totais de contra partida para posteriormente checar se o total de contra partida bate com o lançamento oposto

Dim objContraPartida As ClassContraPartida
Dim iAchou As Integer
Dim lErro As Long

On Error GoTo Erro_Armazena_Contra_Partida

    For Each objContraPartida In colContraPartida
    
        If objContraPartida.iSeqContraPartida = objLancamento_Detalhe.iSeqContraPartida Then
        
            objContraPartida.dValorContraPartida = objContraPartida.dValorContraPartida - objLancamento_Detalhe.dValor
            iAchou = 1
            Exit For
            
        End If
        
    Next
    
    If iAchou = 0 Then
            
        Set objContraPartida = New ClassContraPartida
        
        objContraPartida.iSeqContraPartida = objLancamento_Detalhe.iSeqContraPartida
        objContraPartida.dValorContraPartida = -objLancamento_Detalhe.dValor
                
        colContraPartida.Add objContraPartida
    
    End If

    Armazena_Contra_Partida = SUCESSO
    
    Exit Function

Erro_Armazena_Contra_Partida:

    Armazena_Contra_Partida = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162290)

    End Select
    
    Exit Function
    
End Function

Private Function Testa_Contra_Partida(colLancamento_Detalhe As Collection, colContraPartida As Collection) As Long
'checa se o total de contra partida bate com o lançamento correspondente

Dim objContraPartida As ClassContraPartida
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim iAchou As Integer
Dim lErro As Long

On Error GoTo Erro_Testa_Contra_Partida

    For Each objContraPartida In colContraPartida
        
        For Each objLancamento_Detalhe In colLancamento_Detalhe
        
            If objContraPartida.iSeqContraPartida = objLancamento_Detalhe.iSeq Then
            
                iAchou = 1
                If objContraPartida.dValorContraPartida <> objLancamento_Detalhe.dValor Then Error 20616
                Exit For
                
            End If
            
        Next
        
        'se não achou o lancamento oposto da contra-partida
        If iAchou = 0 Then Error 20617
        
        iAchou = 0
        
    Next
        
    Testa_Contra_Partida = SUCESSO
    
    Exit Function

Erro_Testa_Contra_Partida:

    Testa_Contra_Partida = Err
    
    Select Case Err
    
        Case 20616
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_CONTRA_PARTIDA_VALOR", Err, objLancamento_Detalhe.iSeq, Abs(objLancamento_Detalhe.dValor), Abs(objContraPartida.dValorContraPartida))
        
        Case 20617
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_CONTRA_PARTIDA_INEXISTENTE", Err, objContraPartida.iSeqContraPartida)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162291)

    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_LANCAMENTOS
    Set Form_Load_Ocx = Me
    Caption = "Lançamentos em Lote"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Lancamentos"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Lote Then
            Call Label2_Click
        ElseIf Me.ActiveControl Is Documento Then
            Call Label3_Click
        ElseIf Me.ActiveControl Is DocAuto Then
            Call Label6_Click
        ElseIf Me.ActiveControl Is Conta Then
            Call BotaoConta_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcl_Click
        ElseIf Me.ActiveControl Is Historico Then
            Call BotaoHist_Click
        End If
    
    End If

End Sub


Private Sub CclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclLabel, Source, X, Y)
End Sub

Private Sub CclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclLabel, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub ContaDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaDescricao, Source, X, Y)
End Sub

Private Sub ContaDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaDescricao, Button, Shift, X, Y)
End Sub

Private Sub CclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclDescricao, Source, X, Y)
End Sub

Private Sub CclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclDescricao, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub LabelDocAuto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDocAuto, Source, X, Y)
End Sub

Private Sub LabelDocAuto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDocAuto, Button, Shift, X, Y)
End Sub

Private Sub LabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCcl, Source, X, Y)
End Sub

Private Sub LabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCcl, Button, Shift, X, Y)
End Sub

Private Sub LabelHistoricos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelHistoricos, Source, X, Y)
End Sub

Private Sub LabelHistoricos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelHistoricos, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
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

Private Sub LabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotais, Source, X, Y)
End Sub

Private Sub LabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotais, Button, Shift, X, Y)
End Sub

Private Sub TotalDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalDebito, Source, X, Y)
End Sub

Private Sub TotalDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalDebito, Button, Shift, X, Y)
End Sub

Private Sub TotalCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalCredito, Source, X, Y)
End Sub

Private Sub TotalCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalCredito, Button, Shift, X, Y)
End Sub

Private Sub Periodo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Periodo, Source, X, Y)
End Sub

Private Sub Periodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Periodo, Button, Shift, X, Y)
End Sub

Private Sub Exercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Exercicio, Source, X, Y)
End Sub

Private Sub Exercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Exercicio, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Origem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Origem, Source, X, Y)
End Sub

Private Sub Origem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Origem, Button, Shift, X, Y)
End Sub

Private Sub LabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContas, Source, X, Y)
End Sub

Private Sub LabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContas, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

'#####################################################
'Inserido por Wagner 30/09/05
Private Function Exibe_Dados() As Long

Dim lErro As Long
Dim sConta As String
Dim sCcl As String

On Error GoTo Erro_Exibe_Dados

    If GridLancamentos.Row > 0 Then

        sConta = GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Conta_Col)
    
        If Len(sConta) > 0 Then
            Call Conta_Exibe_Descricao(sConta)
        Else
            ContaDescricao = ""
        End If
        
        If iGrid_Ccl_Col <> 999 Then

            sCcl = GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Ccl_Col)
            
            If Len(sCcl) > 0 Then
                Call Ccl_Exibe_Descricao(sCcl)
            Else
                CclDescricao = ""
            End If
            
        End If
        
    End If
        
    Exibe_Dados = SUCESSO
    
    Exit Function

Erro_Exibe_Dados:

    Exibe_Dados = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162292)

    End Select
    
    Exit Function

End Function
'######################################################

Private Function Saida_Celula_Gerencial(objGridInt As AdmGrid) As Long
'faz a critica da celula Gerencial do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Gerencial

    Set objGridInt.objControle = Gerencial

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 188076

    Saida_Celula_Gerencial = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Gerencial:

    Saida_Celula_Gerencial = gErr
    
    Select Case gErr
    
        Case 188076
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188077)
        
    End Select

    Exit Function

End Function

Private Sub BotaoConta_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Conta.Text) > 0 Then objPlanoConta.sConta = Conta.Text
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("PlanoContaLista", colSelecao, objPlanoConta, objEventoConta)

End Sub

Private Sub objEventoConta_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sConta As String
Dim sContaEnxuta As String
Dim objHistPadrao As New ClassHistPadrao

On Error GoTo Erro_objEventoConta_evSelecao
    
    If GridLancamentos.Col = iGrid_Conta_Col Then

        Set objPlanoConta = obj1
        
        sConta = objPlanoConta.sConta
        
        'le a conta
        lErro = CF("PlanoConta_Le_Conta1", sConta, objPlanoConta)
        If lErro <> SUCESSO And lErro <> 6030 Then gError 197910
        
        If objPlanoConta.iAtivo <> CONTA_ATIVA Then gError 197911
        
        If objPlanoConta.iTipoConta <> CONTA_ANALITICA Then gError 197912
        
        sContaEnxuta = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 197913

        Conta.PromptInclude = False
        Conta.Text = sContaEnxuta
        Conta.PromptInclude = True

        GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col) = Conta.Text

        If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.iLinhasExistentes Then
            objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
        End If

        ContaDescricao.Caption = objPlanoConta.sDescConta
        
        'Se a Conta possui um Histórico Padrão "default" coloca na tela
        If Len(Trim(GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Historico_Col))) = 0 And objPlanoConta.iHistPadrao <> 0 Then
                        
            objHistPadrao.iHistPadrao = objPlanoConta.iHistPadrao
                        
            'le os dados do historico
            lErro = CF("HistPadrao_Le", objHistPadrao)
            If lErro <> SUCESSO And lErro <> 5446 Then gError 197914
                                    
            If lErro = SUCESSO Then
            
                GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Historico_Col) = objHistPadrao.sDescHistPadrao
                
            End If
                        
        End If

    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoConta_evSelecao:

    Select Case gErr
    
        Case 197910, 197914
    
        Case 197911
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_INATIVA", gErr, sConta)
        
        Case 197912
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_ANALITICA", gErr, sConta)
    
        Case 197913
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, sConta)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197915)
        
    End Select

    Exit Sub

End Sub

Private Sub BotaoCcl_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Ccl.Text) > 0 Then objCcl.sCcl = Ccl.Text
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objCcl As ClassCcl
Dim sConta As String
Dim sCclEnxuta As String

On Error GoTo Erro_objEventoCcl_evSelecao
    
    If GridLancamentos.Col = iGrid_Ccl_Col Then

        Set objCcl = obj1

        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO And lErro <> 5599 Then gError 197916

        If objCcl.iTipoCcl <> CCL_ANALITICA Then gError 197917
        
        If objCcl.iAtivo = 0 Then gError 197918
        
        sCclEnxuta = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclEnxuta)
        If lErro <> SUCESSO Then gError 197919

        Ccl.PromptInclude = False
        Ccl.Text = sCclEnxuta
        Ccl.PromptInclude = True

        GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col) = Ccl.Text

        If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.iLinhasExistentes Then
            objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
        End If

        CclDescricao.Caption = objCcl.sDescCcl

    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoCcl_evSelecao:

    Select Case gErr
    
        Case 197916

        Case 197917
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_ANALITICA1", gErr, objCcl.sCcl)
  
        Case 197918
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_INATIVO", gErr, objCcl.sCcl)

        Case 197919
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCcl)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197920)
        
    End Select

    Exit Sub

End Sub

Private Sub BotaoHist_Click()

Dim colSelecao As Collection
Dim objHistPadrao As New ClassHistPadrao

    Call Chama_Tela("HistPadraoLista", colSelecao, objHistPadrao, objEventoHist)

End Sub

Private Sub objEventoHist_evSelecao(obj1 As Object)


Dim objHistPadrao As ClassHistPadrao

On Error GoTo Erro_objEventoHist_evSelecao

    If GridLancamentos.Col = iGrid_Historico_Col Then

        Set objHistPadrao = obj1

        GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col) = objHistPadrao.sDescHistPadrao
        Historico.Text = objHistPadrao.sDescHistPadrao

        If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.iLinhasExistentes Then
            objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
        End If

    End If

    Me.Show
    
    Exit Sub

Erro_objEventoHist_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197921)

    End Select

    Exit Sub

End Sub

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 10

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub

