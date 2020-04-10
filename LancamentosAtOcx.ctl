VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl LancamentosAtOcx 
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11835
   KeyPreview      =   -1  'True
   ScaleHeight     =   6900
   ScaleMode       =   0  'User
   ScaleWidth      =   11840
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
      Left            =   9630
      TabIndex        =   48
      Top             =   4860
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
      Left            =   9630
      TabIndex        =   47
      Top             =   5475
      Width           =   1605
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
      Left            =   9615
      TabIndex        =   46
      Top             =   6090
      Width           =   1605
   End
   Begin VB.CheckBox Gerencial 
      Height          =   210
      Left            =   4875
      TabIndex        =   45
      Tag             =   "1"
      Top             =   2655
      Width           =   870
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   4800
      Picture         =   "LancamentosAtOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   135
      Width           =   300
   End
   Begin VB.PictureBox Picture3 
      Height          =   540
      Left            =   9810
      ScaleHeight     =   480
      ScaleWidth      =   1875
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   165
      Width           =   1935
      Begin VB.CommandButton BotaoFechar 
         Height          =   330
         Left            =   1410
         Picture         =   "LancamentosAtOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   330
         Left            =   960
         Picture         =   "LancamentosAtOcx.ctx":0268
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   330
         Left            =   510
         Picture         =   "LancamentosAtOcx.ctx":079A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   330
         Left            =   60
         Picture         =   "LancamentosAtOcx.ctx":0CCC
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   75
         Width           =   390
      End
   End
   Begin VB.ListBox ListHistoricos 
      Height          =   5520
      Left            =   9255
      TabIndex        =   12
      Top             =   1290
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.ListBox ListDocAuto 
      Height          =   5520
      Left            =   9255
      TabIndex        =   15
      Top             =   1305
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.TextBox Historico 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3600
      MaxLength       =   150
      TabIndex        =   8
      Top             =   2085
      Width           =   4965
   End
   Begin VB.Frame Frame2 
      Caption         =   "Documento Automático"
      Height          =   750
      Left            =   120
      TabIndex        =   23
      Top             =   6060
      Width           =   6315
      Begin VB.CommandButton BotaoAplicar 
         Height          =   510
         Left            =   3960
         Picture         =   "LancamentosAtOcx.ctx":0E26
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   180
         Width           =   1095
      End
      Begin MSMask.MaskEdBox DocAuto 
         Height          =   285
         Left            =   2190
         TabIndex        =   10
         Top             =   330
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
         Left            =   1440
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   330
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Descrição do Elemento Selecionado"
      Height          =   1050
      Left            =   135
      TabIndex        =   22
      Top             =   4950
      Width           =   6315
      Begin VB.Label CclDescricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1845
         TabIndex        =   25
         Top             =   645
         Visible         =   0   'False
         Width           =   3720
      End
      Begin VB.Label ContaDescricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1845
         TabIndex        =   26
         Top             =   285
         Width           =   3720
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
         Top             =   315
         Width           =   570
      End
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
         TabIndex        =   28
         Top             =   660
         Visible         =   0   'False
         Width           =   1440
      End
   End
   Begin MSMask.MaskEdBox SeqContraPartida 
      Height          =   225
      Left            =   4755
      TabIndex        =   7
      Top             =   1845
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
      Left            =   5265
      TabIndex        =   6
      Top             =   1335
      Width           =   825
      _ExtentX        =   1455
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
      Left            =   4140
      TabIndex        =   5
      Top             =   1455
      Width           =   825
      _ExtentX        =   1455
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
      Left            =   1590
      TabIndex        =   4
      Top             =   2445
      Width           =   1005
      _ExtentX        =   1773
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
      Left            =   450
      TabIndex        =   3
      Top             =   1830
      Width           =   1410
      _ExtentX        =   2487
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
      Left            =   9255
      TabIndex        =   14
      Top             =   1305
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   9737
      _Version        =   393217
      Indentation     =   511
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   285
      Left            =   2010
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   510
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSFlexGridLib.MSFlexGrid GridLancamentos 
      Height          =   1860
      Left            =   105
      TabIndex        =   9
      Top             =   1215
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   3281
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin MSComctlLib.TreeView TvwCcls 
      Height          =   5520
      Left            =   9255
      TabIndex        =   13
      Top             =   1305
      Visible         =   0   'False
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   9737
      _Version        =   393217
      Indentation     =   511
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSMask.MaskEdBox Documento 
      Height          =   285
      Left            =   3585
      TabIndex        =   0
      Top             =   135
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSMask.MaskEdBox Data 
      Height          =   285
      Left            =   885
      TabIndex        =   2
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
      Left            =   9330
      TabIndex        =   29
      Top             =   1095
      Width           =   2235
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
      Left            =   9330
      TabIndex        =   30
      Top             =   1095
      Width           =   2400
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
      Left            =   2655
      TabIndex        =   31
      Top             =   540
      Width           =   870
   End
   Begin VB.Label TotalCredito 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2190
      TabIndex        =   32
      Top             =   4500
      Width           =   1155
   End
   Begin VB.Label TotalDebito 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3435
      TabIndex        =   33
      Top             =   4515
      Width           =   1155
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
      Left            =   1380
      TabIndex        =   34
      Top             =   4560
      Width           =   705
   End
   Begin VB.Label Label4 
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
      Left            =   330
      TabIndex        =   35
      Top             =   540
      Width           =   480
   End
   Begin VB.Label DocumentoLabel 
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
      Left            =   2520
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   36
      Top             =   165
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   150
      TabIndex        =   37
      Top             =   150
      Width           =   660
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
      Left            =   5010
      TabIndex        =   38
      Top             =   540
      Width           =   735
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
      Left            =   9330
      TabIndex        =   39
      Top             =   1095
      Width           =   2370
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
      Left            =   9330
      TabIndex        =   40
      Top             =   1095
      Width           =   2460
   End
   Begin VB.Label Origem 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Contabilidade"
      Height          =   285
      Left            =   900
      TabIndex        =   41
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label Periodo 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5820
      TabIndex        =   42
      Top             =   510
      Width           =   1185
   End
   Begin VB.Label Exercicio 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3585
      TabIndex        =   43
      Top             =   510
      Width           =   1185
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
      TabIndex        =   44
      Top             =   990
      Width           =   1140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   -120.051
      X2              =   11779.97
      Y1              =   885
      Y2              =   885
   End
End
Attribute VB_Name = "LancamentosAtOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTLancamentosAt
Attribute objCT.VB_VarHelpID = -1

Public Sub menuRateio_Click()
     Call objCT.menuRateio_Click
End Sub

Public Sub menulimpar_Click()
     Call objCT.menulimpar_Click
End Sub

Private Sub BotaoCcl_Click()
     Call objCT.BotaoCcl_Click
End Sub

Private Sub BotaoConta_Click()
     Call objCT.BotaoConta_Click
End Sub

Private Sub BotaoHist_Click()
     Call objCT.BotaoHist_Click
End Sub

Private Sub BotaoProxNum_Click()
     Call objCT.BotaoProxNum_Click
End Sub

Private Sub BotaoAplicar_Click()
     Call objCT.BotaoAplicar_Click
End Sub

Private Sub Ccl_Change()
     Call objCT.Ccl_Change
End Sub

Private Sub Conta_Change()
     Call objCT.Conta_Change
End Sub

Private Sub Conta_GotFocus()
     Call objCT.Conta_GotFocus
End Sub

Private Sub Conta_KeyPress(KeyAscii As Integer)
     Call objCT.Conta_KeyPress(KeyAscii)
End Sub

Private Sub Conta_Validate(Cancel As Boolean)
     Call objCT.Conta_Validate(Cancel)
End Sub

Private Sub Ccl_GotFocus()
     Call objCT.Ccl_GotFocus
End Sub

Private Sub Data_GotFocus()
     Call objCT.Data_GotFocus
End Sub

Private Sub DocAuto_GotFocus()
     Call objCT.DocAuto_GotFocus
End Sub

Private Sub Ccl_KeyPress(KeyAscii As Integer)
     Call objCT.Ccl_KeyPress(KeyAscii)
End Sub

Private Sub Ccl_Validate(Cancel As Boolean)
     Call objCT.Ccl_Validate(Cancel)
End Sub

Private Sub Credito_Change()
     Call objCT.Credito_Change
End Sub

Private Sub Credito_GotFocus()
     Call objCT.Credito_GotFocus
End Sub

Private Sub Credito_KeyPress(KeyAscii As Integer)
     Call objCT.Credito_KeyPress(KeyAscii)
End Sub

Private Sub Credito_Validate(Cancel As Boolean)
     Call objCT.Credito_Validate(Cancel)
End Sub

Private Sub Documento_GotFocus()
     Call objCT.Documento_GotFocus
End Sub

Private Sub DocumentoLabel_Click()
    Call objCT.DocumentoLabel_Click
End Sub

Private Sub SeqContraPartida_Change()
     Call objCT.SeqContraPartida_Change
End Sub

Private Sub SeqContraPartida_GotFocus()
     Call objCT.SeqContraPartida_GotFocus
End Sub

Private Sub SeqContraPartida_KeyPress(KeyAscii As Integer)
     Call objCT.SeqContraPartida_KeyPress(KeyAscii)
End Sub

Private Sub SeqContraPartida_Validate(Cancel As Boolean)
     Call objCT.SeqContraPartida_Validate(Cancel)
End Sub

Private Sub Data_Change()
     Call objCT.Data_Change
End Sub

Private Sub Debito_Change()
     Call objCT.Debito_Change
End Sub

Private Sub Debito_GotFocus()
     Call objCT.Debito_GotFocus
End Sub

Private Sub Debito_KeyPress(KeyAscii As Integer)
     Call objCT.Debito_KeyPress(KeyAscii)
End Sub

Private Sub Debito_Validate(Cancel As Boolean)
     Call objCT.Debito_Validate(Cancel)
End Sub

Private Sub DocAuto_Change()
     Call objCT.DocAuto_Change
End Sub

Private Sub Documento_Change()
     Call objCT.Documento_Change
End Sub

Private Sub Exercicio_Change()
     Call objCT.Exercicio_Change
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub GridLancamentos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Faz com que apareca um PopupMenu o botao direito do mouse acionado sobre o grid

     Call objCT.GridLancamentos_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Historico_Change()
     Call objCT.Historico_Change
End Sub

Private Sub Historico_GotFocus()
     Call objCT.Historico_GotFocus
End Sub

Private Sub Historico_KeyPress(KeyAscii As Integer)
     Call objCT.Historico_KeyPress(KeyAscii)
End Sub

Private Sub Historico_Validate(Cancel As Boolean)
     Call objCT.Historico_Validate(Cancel)
End Sub

Private Sub GridLancamentos_Click()
     Call objCT.GridLancamentos_Click
End Sub

Private Sub GridLancamentos_GotFocus()
     Call objCT.GridLancamentos_GotFocus
End Sub

Private Sub GridLancamentos_EnterCell()
     Call objCT.GridLancamentos_EnterCell
End Sub

Private Sub GridLancamentos_LeaveCell()
     Call objCT.GridLancamentos_LeaveCell
End Sub

Private Sub GridLancamentos_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridLancamentos_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridLancamentos_KeyPress(KeyAscii As Integer)
     Call objCT.GridLancamentos_KeyPress(KeyAscii)
End Sub

Private Sub GridLancamentos_Validate(Cancel As Boolean)
     Call objCT.GridLancamentos_Validate(Cancel)
End Sub

Private Sub GridLancamentos_RowColChange()
     Call objCT.GridLancamentos_RowColChange
End Sub

Private Sub GridLancamentos_Scroll()
     Call objCT.GridLancamentos_Scroll
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Function Trata_Parametros() As Long
     Trata_Parametros = objCT.Trata_Parametros()
End Function

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub Label6_Click()
     Call objCT.Label6_Click
End Sub

Private Sub Data_Validate(Cancel As Boolean)
     Call objCT.Data_Validate(Cancel)
End Sub

Private Sub TvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.TvwCcls_NodeClick(Node)
End Sub

Private Sub ListDocAuto_DblClick()
     Call objCT.ListDocAuto_DblClick
End Sub

Private Sub ListHistoricos_DblClick()
     Call objCT.ListHistoricos_DblClick
End Sub

Private Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)
     Call objCT.TvwContas_Expand(objNode)
End Sub

Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.TvwContas_NodeClick(Node)
End Sub

Private Sub UpDown1_DownClick()
     Call objCT.UpDown1_DownClick
End Sub

Private Sub UpDown1_UpClick()
     Call objCT.UpDown1_UpClick
End Sub

Private Sub BotaoImprimir_Click()
     Call objCT.BotaoImprimir_Click
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

Private Sub UserControl_Initialize()
    Set objCT = New CTLancamentosAt
    Set objCT.objUserControl = Me
End Sub

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

Public Sub Popup_Menu(Menu As Object)

    UserControl.PopupMenu Menu
    
End Sub




Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub CclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclDescricao, Source, X, Y)
End Sub

Private Sub CclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclDescricao, Button, Shift, X, Y)
End Sub

Private Sub ContaDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaDescricao, Source, X, Y)
End Sub

Private Sub ContaDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaDescricao, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub CclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclLabel, Source, X, Y)
End Sub

Private Sub CclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelHistoricos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelHistoricos, Source, X, Y)
End Sub

Private Sub LabelHistoricos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelHistoricos, Button, Shift, X, Y)
End Sub

Private Sub LabelDocAuto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDocAuto, Source, X, Y)
End Sub

Private Sub LabelDocAuto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDocAuto, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub TotalCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalCredito, Source, X, Y)
End Sub

Private Sub TotalCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalCredito, Button, Shift, X, Y)
End Sub

Private Sub TotalDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalDebito, Source, X, Y)
End Sub

Private Sub TotalDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalDebito, Button, Shift, X, Y)
End Sub

Private Sub LabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotais, Source, X, Y)
End Sub

Private Sub LabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotais, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub LabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCcl, Source, X, Y)
End Sub

Private Sub LabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCcl, Button, Shift, X, Y)
End Sub

Private Sub LabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContas, Source, X, Y)
End Sub

Private Sub LabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContas, Button, Shift, X, Y)
End Sub

Private Sub Origem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Origem, Source, X, Y)
End Sub

Private Sub Origem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Origem, Button, Shift, X, Y)
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

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Gerencial_GotFocus()

    Call objCT.Gerencial_GotFocus

End Sub

Private Sub Gerencial_KeyPress(KeyAscii As Integer)

    Call objCT.Gerencial_KeyPress(KeyAscii)

End Sub

Private Sub Gerencial_Validate(Cancel As Boolean)

    Call objCT.Gerencial_Validate(Cancel)

End Sub

Private Sub DocumentoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DocumentoLabel, Source, X, Y)
End Sub

Private Sub DocumentoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DocumentoLabel, Button, Shift, X, Y)
End Sub


