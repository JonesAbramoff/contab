VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl PadraoContabOcx 
   ClientHeight    =   6165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9465
   KeyPreview      =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   9465
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1875
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   2295
      Visible         =   0   'False
      Width           =   9105
      Begin VB.CheckBox GerencialCusto 
         Height          =   210
         Left            =   7230
         TabIndex        =   44
         Tag             =   "1"
         Top             =   1350
         Width           =   870
      End
      Begin VB.TextBox HistoricoCusto 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4155
         MaxLength       =   255
         TabIndex        =   18
         Top             =   885
         Width           =   2475
      End
      Begin VB.TextBox ContaCusto 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   360
         MaxLength       =   255
         TabIndex        =   14
         Top             =   1200
         Width           =   2475
      End
      Begin VB.TextBox CclCusto 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2160
         MaxLength       =   255
         TabIndex        =   15
         Top             =   1440
         Width           =   2475
      End
      Begin VB.TextBox CreditoCusto 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2160
         MaxLength       =   255
         TabIndex        =   16
         Top             =   1200
         Width           =   2475
      End
      Begin VB.TextBox DebitoCusto 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4185
         MaxLength       =   255
         TabIndex        =   17
         Top             =   1200
         Width           =   2475
      End
      Begin VB.CheckBox AglutinaCusto 
         Height          =   210
         Left            =   6090
         TabIndex        =   20
         Top             =   1350
         Width           =   870
      End
      Begin MSMask.MaskEdBox ProdutoCusto 
         Height          =   225
         Left            =   6690
         TabIndex        =   19
         Top             =   915
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridCusto 
         Height          =   1620
         Left            =   0
         TabIndex        =   21
         Top             =   120
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   2858
         _Version        =   393216
         Rows            =   11
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   0   'False
         FocusRect       =   2
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1875
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   2325
      Width           =   9105
      Begin VB.CheckBox Gerencial 
         Height          =   210
         Left            =   5625
         TabIndex        =   43
         Tag             =   "1"
         Top             =   1185
         Width           =   870
      End
      Begin VB.TextBox Debito 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4200
         MaxLength       =   255
         TabIndex        =   9
         Top             =   720
         Width           =   2475
      End
      Begin VB.TextBox Credito 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1920
         MaxLength       =   255
         TabIndex        =   8
         Top             =   840
         Width           =   2475
      End
      Begin VB.TextBox Ccl 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2175
         MaxLength       =   255
         TabIndex        =   7
         Top             =   360
         Width           =   2475
      End
      Begin VB.TextBox Conta 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   240
         MaxLength       =   255
         TabIndex        =   6
         Top             =   600
         Width           =   2475
      End
      Begin VB.TextBox Historico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   6360
         MaxLength       =   255
         TabIndex        =   10
         Top             =   840
         Width           =   2475
      End
      Begin VB.CheckBox Aglutina 
         Height          =   210
         Left            =   7440
         TabIndex        =   11
         Top             =   1200
         Width           =   870
      End
      Begin MSFlexGridLib.MSFlexGrid GridLancamentos 
         Height          =   1665
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   2937
         _Version        =   393216
         Rows            =   11
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   0   'False
         FocusRect       =   2
      End
   End
   Begin VB.CheckBox Obrigatoriedade 
      Caption         =   "A contabilização desta transação é obrigatória"
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
      Left            =   1320
      TabIndex        =   42
      Top             =   1050
      Width           =   4695
   End
   Begin VB.TextBox Descricao 
      BackColor       =   &H8000000F&
      Height          =   540
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   41
      Top             =   5490
      Width           =   8505
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7260
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "PadraoContabOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "PadraoContabOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "PadraoContabOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "PadraoContabOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox Modulo 
      Height          =   315
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3630
   End
   Begin VB.ComboBox Transacao 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   645
      Width           =   5745
   End
   Begin VB.CommandButton BotaoPlanoConta 
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
      Height          =   330
      Left            =   435
      TabIndex        =   22
      Top             =   4410
      Width           =   1875
   End
   Begin VB.CommandButton BotaoCcl 
      Caption         =   "Centros de Custo/Lucro"
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
      Left            =   4620
      TabIndex        =   24
      Top             =   4410
      Width           =   2325
   End
   Begin VB.CommandButton BotaoHistorico 
      Caption         =   "Históricos Padrão"
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
      Left            =   2490
      TabIndex        =   23
      Top             =   4410
      Width           =   1965
   End
   Begin VB.ComboBox Modelo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1350
      TabIndex        =   2
      Text            =   "Modelo"
      Top             =   1470
      Width           =   3630
   End
   Begin VB.ComboBox Operadores 
      Height          =   315
      Left            =   7830
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   5010
      Width           =   1050
   End
   Begin VB.ComboBox Funcoes 
      Height          =   315
      Left            =   4080
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   5010
      Width           =   3555
   End
   Begin VB.ComboBox Mnemonicos 
      Enabled         =   0   'False
      Height          =   315
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   5010
      Width           =   3555
   End
   Begin VB.CommandButton BotaoProduto 
      Caption         =   "Produto"
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
      Left            =   7110
      TabIndex        =   25
      Top             =   4410
      Width           =   1770
   End
   Begin VB.CheckBox Checkbox_Verifica_Sintaxe 
      Caption         =   "Verifica Sintaxe ao Sair da Célula (F5)"
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
      Left            =   5730
      TabIndex        =   4
      Top             =   1935
      Value           =   1  'Checked
      Width           =   3600
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2295
      Left            =   120
      TabIndex        =   34
      Top             =   1935
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   4048
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Lançamentos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Custo"
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
   Begin VB.CheckBox Padrao 
      Caption         =   "Padrão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5280
      TabIndex        =   3
      Top             =   1470
      Width           =   990
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
      Left            =   555
      TabIndex        =   40
      Top             =   165
      Width           =   690
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
      Left            =   270
      TabIndex        =   39
      Top             =   690
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Modelo:"
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
      Left            =   600
      TabIndex        =   38
      Top             =   1470
      Width           =   690
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Campos:"
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
      Left            =   390
      TabIndex        =   37
      Top             =   4770
      Width           =   735
   End
   Begin VB.Label Label5 
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
      Left            =   4080
      TabIndex        =   36
      Top             =   4770
      Width           =   795
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
      Left            =   7830
      TabIndex        =   35
      Top             =   4770
      Width           =   1050
   End
End
Attribute VB_Name = "PadraoContabOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTPadraoContab
Attribute objCT.VB_VarHelpID = -1

Private Sub BotaoCcl_Click()
     Call objCT.BotaoCcl_Click
End Sub

Private Sub BotaoHistorico_Click()
     Call objCT.BotaoHistorico_Click
End Sub

Private Sub BotaoPlanoConta_Click()
     Call objCT.BotaoPlanoConta_Click
End Sub

Private Sub BotaoProduto_Click()
     Call objCT.BotaoProduto_Click
End Sub

Private Sub Ccl_Change()
     Call objCT.Ccl_Change
End Sub

Private Sub CclCusto_Change()
     Call objCT.CclCusto_Change
End Sub

Private Sub CclCusto_GotFocus()
     Call objCT.CclCusto_GotFocus
End Sub

Private Sub CclCusto_KeyPress(KeyAscii As Integer)
     Call objCT.CclCusto_KeyPress(KeyAscii)
End Sub

Private Sub CclCusto_Validate(Cancel As Boolean)
     Call objCT.CclCusto_Validate(Cancel)
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

Private Sub Ccl_KeyPress(KeyAscii As Integer)
     Call objCT.Ccl_KeyPress(KeyAscii)
End Sub

Private Sub Ccl_Validate(Cancel As Boolean)
     Call objCT.Ccl_Validate(Cancel)
End Sub

Private Sub ContaCusto_Change()
     Call objCT.ContaCusto_Change
End Sub

Private Sub ContaCusto_GotFocus()
     Call objCT.ContaCusto_GotFocus
End Sub

Private Sub ContaCusto_KeyPress(KeyAscii As Integer)
     Call objCT.ContaCusto_KeyPress(KeyAscii)
End Sub

Private Sub ContaCusto_Validate(Cancel As Boolean)
     Call objCT.ContaCusto_Validate(Cancel)
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

Private Sub CreditoCusto_Change()
     Call objCT.CreditoCusto_Change
End Sub

Private Sub CreditoCusto_GotFocus()
     Call objCT.CreditoCusto_GotFocus
End Sub

Private Sub CreditoCusto_KeyPress(KeyAscii As Integer)
     Call objCT.CreditoCusto_KeyPress(KeyAscii)
End Sub

Private Sub CreditoCusto_Validate(Cancel As Boolean)
     Call objCT.CreditoCusto_Validate(Cancel)
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

Private Sub DebitoCusto_Change()
     Call objCT.DebitoCusto_Change
End Sub

Private Sub DebitoCusto_GotFocus()
     Call objCT.DebitoCusto_GotFocus
End Sub

Private Sub DebitoCusto_KeyPress(KeyAscii As Integer)
     Call objCT.DebitoCusto_KeyPress(KeyAscii)
End Sub

Private Sub DebitoCusto_Validate(Cancel As Boolean)
     Call objCT.DebitoCusto_Validate(Cancel)
End Sub

Private Sub Aglutina_GotFocus()
     Call objCT.Aglutina_GotFocus
End Sub

Private Sub Aglutina_KeyPress(KeyAscii As Integer)
     Call objCT.Aglutina_KeyPress(KeyAscii)
End Sub

Private Sub Aglutina_Validate(Cancel As Boolean)
     Call objCT.Aglutina_Validate(Cancel)
End Sub

Private Sub AglutinaCusto_GotFocus()
     Call objCT.AglutinaCusto_GotFocus
End Sub

Private Sub AglutinaCusto_KeyPress(KeyAscii As Integer)
     Call objCT.AglutinaCusto_KeyPress(KeyAscii)
End Sub

Private Sub AglutinaCusto_Validate(Cancel As Boolean)
     Call objCT.AglutinaCusto_Validate(Cancel)
End Sub

Private Sub Obrigatoriedade_Click()
    Call objCT.Obrigatoriedade_Click
End Sub

Private Sub ProdutoCusto_Change()
     Call objCT.ProdutoCusto_Change
End Sub

Private Sub ProdutoCusto_GotFocus()
     Call objCT.ProdutoCusto_GotFocus
End Sub

Private Sub ProdutoCusto_KeyPress(KeyAscii As Integer)
     Call objCT.ProdutoCusto_KeyPress(KeyAscii)
End Sub

Private Sub ProdutoCusto_Validate(Cancel As Boolean)
     Call objCT.ProdutoCusto_Validate(Cancel)
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub Funcoes_Click()
     Call objCT.Funcoes_Click
End Sub

Private Sub GridCusto_Click()
     Call objCT.GridCusto_Click
End Sub

Private Sub GridCusto_EnterCell()
     Call objCT.GridCusto_EnterCell
End Sub

Private Sub GridCusto_GotFocus()
     Call objCT.GridCusto_GotFocus
End Sub

Private Sub GridCusto_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridCusto_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridCusto_KeyPress(KeyAscii As Integer)
     Call objCT.GridCusto_KeyPress(KeyAscii)
End Sub

Private Sub GridCusto_LeaveCell()
     Call objCT.GridCusto_LeaveCell
End Sub

Private Sub GridCusto_Validate(Cancel As Boolean)
     Call objCT.GridCusto_Validate(Cancel)
End Sub

Private Sub GridCusto_RowColChange()
     Call objCT.GridCusto_RowColChange
End Sub

Private Sub GridCusto_Scroll()
     Call objCT.GridCusto_Scroll
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

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub HistoricoCusto_Change()
     Call objCT.HistoricoCusto_Change
End Sub

Private Sub HistoricoCusto_GotFocus()
     Call objCT.HistoricoCusto_GotFocus
End Sub

Private Sub HistoricoCusto_KeyPress(KeyAscii As Integer)
     Call objCT.HistoricoCusto_KeyPress(KeyAscii)
End Sub

Private Sub HistoricoCusto_Validate(Cancel As Boolean)
     Call objCT.HistoricoCusto_Validate(Cancel)
End Sub

Private Sub Mnemonicos_Click()
     Call objCT.Mnemonicos_Click
End Sub

Private Sub Modelo_Change()
     Call objCT.Modelo_Change
End Sub

Private Sub Modelo_Click()
     Call objCT.Modelo_Click
End Sub

Private Sub Modulo_Click()
     Call objCT.Modulo_Click
End Sub

Private Sub Operadores_Click()
     Call objCT.Operadores_Click
End Sub

Private Sub Padrao_Click()
     Call objCT.Padrao_Click
End Sub

Private Sub TabStrip1_Click()
     Call objCT.TabStrip1_Click
End Sub

Private Sub Transacao_Click()
     Call objCT.Transacao_Click
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
    Set objCT = New CTPadraoContab
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

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub


Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
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

Private Sub GerencialCusto_GotFocus()
     Call objCT.GerencialCusto_GotFocus
End Sub

Private Sub GerencialCusto_KeyPress(KeyAscii As Integer)
     Call objCT.GerencialCusto_KeyPress(KeyAscii)
End Sub

Private Sub GerencialCusto_Validate(Cancel As Boolean)
     Call objCT.GerencialCusto_Validate(Cancel)
End Sub

