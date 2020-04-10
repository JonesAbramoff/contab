VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.UserControl InventarioLote 
   ClientHeight    =   4980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9375
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4980
   ScaleMode       =   0  'User
   ScaleWidth      =   9360.235
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2160
      Picture         =   "InventarioLote.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Numeração Automática"
      Top             =   195
      Width           =   300
   End
   Begin VB.TextBox Codigo 
      Height          =   300
      Left            =   960
      MaxLength       =   10
      TabIndex        =   1
      Top             =   180
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "objMnemonicoCTBValor.sValor"
      Height          =   3885
      Index           =   1
      Left            =   45
      TabIndex        =   0
      Top             =   1035
      Width           =   9255
      Begin VB.CommandButton BotaoLote 
         Caption         =   "Lote"
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
         Left            =   7680
         TabIndex        =   39
         Top             =   3495
         Width           =   1365
      End
      Begin VB.ComboBox Atualiza 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "InventarioLote.ctx":00EA
         Left            =   7470
         List            =   "InventarioLote.ctx":00F4
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   1860
         Width           =   945
      End
      Begin VB.ComboBox FilialOP 
         Height          =   315
         Left            =   6435
         TabIndex        =   35
         Top             =   540
         Width           =   2160
      End
      Begin MSMask.MaskEdBox LoteProduto 
         Height          =   270
         Left            =   7125
         TabIndex        =   37
         Top             =   210
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.TextBox Etiqueta 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   3825
         MaxLength       =   10
         TabIndex        =   11
         Top             =   1560
         Width           =   1305
      End
      Begin VB.ComboBox UnidadeMed 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1545
         Width           =   660
      End
      Begin VB.TextBox DescricaoItem 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   2445
         MaxLength       =   50
         TabIndex        =   17
         Top             =   2865
         Width           =   2250
      End
      Begin VB.ComboBox Tipo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "InventarioLote.ctx":0107
         Left            =   5235
         List            =   "InventarioLote.ctx":0109
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1860
         Width           =   2115
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
         Height          =   360
         Left            =   5785
         TabIndex        =   24
         Top             =   3495
         Width           =   1815
      End
      Begin VB.CommandButton BotaoProdutos 
         Caption         =   "Produtos"
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
         Left            =   2880
         TabIndex        =   22
         Top             =   3495
         Width           =   1380
      End
      Begin VB.CommandButton BotaoEstoque 
         Caption         =   "Estoque"
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
         Left            =   4340
         TabIndex        =   23
         Top             =   3495
         Width           =   1365
      End
      Begin MSMask.MaskEdBox ContaAjuste 
         Height          =   240
         Left            =   7200
         TabIndex        =   19
         Top             =   1980
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   423
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
      Begin MSMask.MaskEdBox ContaContabil 
         Height          =   240
         Left            =   7155
         TabIndex        =   18
         Top             =   1575
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
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
      Begin MSMask.MaskEdBox ValorDiferenca 
         Height          =   240
         Left            =   5280
         TabIndex        =   16
         Top             =   2835
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox QuantDiferenca 
         Height          =   240
         Left            =   4170
         TabIndex        =   15
         Top             =   2355
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox QuantEstoque 
         Height          =   240
         Left            =   6525
         TabIndex        =   20
         Top             =   2415
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Almoxarifado 
         Height          =   240
         Left            =   2490
         TabIndex        =   10
         Top             =   1560
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   240
         Left            =   210
         TabIndex        =   8
         Top             =   1590
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CustoUnitario 
         Height          =   240
         Left            =   3975
         TabIndex        =   12
         Top             =   1905
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   423
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
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   240
         Left            =   2580
         TabIndex        =   14
         Top             =   2310
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   423
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
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   3030
         Left            =   120
         TabIndex        =   21
         Top             =   315
         Width           =   9060
         _ExtentX        =   15981
         _ExtentY        =   5345
         _Version        =   393216
         Rows            =   21
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Materiais do Inventário"
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
         Left            =   195
         TabIndex        =   30
         Top             =   45
         Width           =   1965
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7065
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   210
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "InventarioLote.ctx":010B
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "InventarioLote.ctx":0265
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "InventarioLote.ctx":03EF
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "InventarioLote.ctx":0921
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComCtl2.UpDown UpDownData 
      Height          =   300
      Left            =   4380
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   195
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   300
      Left            =   3300
      TabIndex        =   3
      Top             =   180
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox AlmoxPadrao 
      Height          =   300
      Left            =   4515
      TabIndex        =   7
      Top             =   675
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Lote 
      Height          =   300
      Left            =   960
      TabIndex        =   5
      Top             =   615
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
      _Version        =   393216
      ClipMode        =   1
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
   Begin MSMask.MaskEdBox Hora 
      Height          =   300
      Left            =   5625
      TabIndex        =   4
      Top             =   225
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
      Caption         =   "Hora:"
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
      Left            =   5085
      TabIndex        =   38
      Top             =   210
      Width           =   480
   End
   Begin VB.Label LoteLabel 
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
      Left            =   435
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   34
      Top             =   645
      Width           =   450
   End
   Begin VB.Label CodigoLabel 
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
      Left            =   255
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   33
      Top             =   210
      Width           =   660
   End
   Begin VB.Label Label2 
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
      Left            =   2745
      TabIndex        =   32
      Top             =   210
      Width           =   480
   End
   Begin VB.Label AlmoxPadraoLabel 
      AutoSize        =   -1  'True
      Caption         =   "Almoxarifado Padrão:"
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
      Left            =   2670
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   31
      Top             =   690
      Width           =   1815
   End
End
Attribute VB_Name = "InventarioLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Public WithEvents objCT As CTInventarioLote
Attribute objCT.VB_VarHelpID = -1

''Property Variables:
'Dim m_Caption As String
'Event Unload()
'
'Public iAlterado As Integer
'Dim iLinhaAntiga As Integer
'Dim iFrameAtual As Integer
'
'Public objGrid As AdmGrid
'Public iGrid_Sequencial_Col As Integer
'Public iGrid_Produto_Col As Integer
'Public iGrid_DescricaoItem_Col As Integer
'Public iGrid_UnidadeMed_Col As Integer
'Public iGrid_Quantidade_Col As Integer
'Public iGrid_Almoxarifado_Col As Integer
'Public iGrid_Etiqueta_Col As Integer
'Public iGrid_CustoUnitario_Col As Integer
'Public iGrid_Tipo_Col As Integer
'Public iGrid_ContaContabil_Col As Integer
'Public iGrid_QuantEstoque_Col As Integer
'Public iGrid_QuantDiferenca_Col As Integer
'Public iGrid_ValorDiferenca_col As Integer
'Public iGrid_ContaAjuste_Col As Integer
'Public iGrid_LoteProduto_Col As Integer
'Public iGrid_FilialOP_Col As Integer
'Public iGrid_Atualiza_Col As Integer
'
'Private WithEvents objEventoCodigo As AdmEvento
'Private WithEvents objEventoLoteInv As AdmEvento
'Private WithEvents objEventoProduto As AdmEvento
'Private WithEvents objEventoEstoque As AdmEvento
'Private WithEvents objEventoAlmoxPadrao As AdmEvento
'Private WithEvents objEventoContaContabil As AdmEvento
'Private WithEvents objEventoRastroLote As AdmEvento
'
''Contante que informa qual é o Mnemonico global
'Const CTAAJUSTEINV As String = "CtaAjusteInv"
'
''Constantes públicas dos tabs
'Private Const TAB_Lancamentos = 1
'Private Const TAB_Contabilizacao = 2

Private Sub BotaoPlanoConta_Click()
     Call objCT.BotaoPlanoConta_Click
End Sub

Private Sub ContaContabil_Change()
     Call objCT.ContaContabil_Change
End Sub

Private Sub ContaAjuste_Change()
     Call objCT.ContaAjuste_Change
End Sub

Private Sub ContaAjuste_GotFocus()
     Call objCT.ContaAjuste_GotFocus
End Sub

Private Sub ContaAjuste_KeyPress(KeyAscii As Integer)
     Call objCT.ContaAjuste_KeyPress(KeyAscii)
End Sub

Private Sub ContaAjuste_Validate(Cancel As Boolean)
     Call objCT.ContaAjuste_Validate(Cancel)
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub LoteProduto_Change()
     Call objCT.LoteProduto_Change
End Sub

Private Sub LoteProduto_GotFocus()
     Call objCT.LoteProduto_GotFocus
End Sub

Private Sub LoteProduto_KeyPress(KeyAscii As Integer)
     Call objCT.LoteProduto_KeyPress(KeyAscii)
End Sub

Private Sub LoteProduto_Validate(Cancel As Boolean)
     Call objCT.LoteProduto_Validate(Cancel)
End Sub

Private Sub Atualiza_Click()
     Call objCT.Atualiza_Click
End Sub

Private Sub Atualiza_GotFocus()
     Call objCT.Atualiza_GotFocus
End Sub

Private Sub Atualiza_KeyPress(KeyAscii As Integer)
     Call objCT.Atualiza_KeyPress(KeyAscii)
End Sub

Private Sub Atualiza_Validate(Cancel As Boolean)
     Call objCT.Atualiza_Validate(Cancel)
End Sub

Private Sub FilialOP_Change()
     Call objCT.FilialOP_Change
End Sub

Private Sub FilialOP_GotFocus()
     Call objCT.FilialOP_GotFocus
End Sub

Private Sub FilialOP_KeyPress(KeyAscii As Integer)
     Call objCT.FilialOP_KeyPress(KeyAscii)
End Sub

Private Sub FilialOP_Validate(Cancel As Boolean)
     Call objCT.FilialOP_Validate(Cancel)
End Sub

Private Sub CodigoLabel_Click()
     Call objCT.CodigoLabel_Click
End Sub

Private Sub Data_GotFocus()
     Call objCT.Data_GotFocus
End Sub

Private Sub GridItens_RowColChange()
     Call objCT.GridItens_RowColChange
End Sub

Private Sub Lote_GotFocus()
     Call objCT.Lote_GotFocus
End Sub


Private Sub Lote_Validate(Cancel As Boolean)
     Call objCT.Lote_Validate(Cancel)
End Sub

Private Sub LoteLabel_Click()
     Call objCT.LoteLabel_Click
End Sub

Private Sub AlmoxPadraoLabel_Click()
     Call objCT.AlmoxPadraoLabel_Click
End Sub


Private Sub BotaoProdutos_Click()
     Call objCT.BotaoProdutos_Click
End Sub

Private Sub BotaoEstoque_Click()
    Call objCT.BotaoEstoque_Click
End Sub

Function Trata_Parametros(Optional objInventario As ClassInventario) As Long
     Trata_Parametros = objCT.Trata_Parametros(objInventario)
End Function

Private Sub AlmoxPadrao_Validate(Cancel As Boolean)
    Call objCT.AlmoxPadrao_Validate(Cancel)
End Sub

Private Sub Data_Validate(Cancel As Boolean)
    Call objCT.Data_Validate(Cancel)
End Sub

'hora
Public Sub Hora_GotFocus()
    Call objCT.Hora_GotFocus
End Sub

'hora
Public Sub Hora_Change()
    Call objCT.Hora_Change
End Sub

'hora
Public Sub Hora_Validate(Cancel As Boolean)
    Call objCT.Hora_Validate(Cancel)
End Sub

Private Sub Tipo_Click()
    Call objCT.Tipo_Click
End Sub

Private Sub UpDownData_DownClick()
    Call objCT.UpDownData_DownClick
End Sub

Private Sub UpDownData_UpClick()
    Call objCT.UpDownData_UpClick
End Sub




Private Sub BotaoGravar_Click()
    Call objCT.BotaoGravar_Click
End Sub


Private Sub BotaoLimpar_Click()
    Call objCT.BotaoLimpar_Click
End Sub

Private Sub BotaoExcluir_Click()
    Call objCT.BotaoExcluir_Click
End Sub

Private Sub BotaoFechar_Click()
    Call objCT.BotaoFechar_Click
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub


Private Sub Almoxarifado_Change()
    Call objCT.Almoxarifado_Change
End Sub

Private Sub AlmoxPadrao_Change()
    Call objCT.AlmoxPadrao_Change
End Sub

Private Sub Codigo_Change()
    Call objCT.Codigo_Change
End Sub

Private Sub Data_Change()
    Call objCT.Data_Change
End Sub

Private Sub CustoUnitario_Change()
    Call objCT.CustoUnitario_Change
End Sub

Private Sub Produto_Change()
    Call objCT.Produto_Change
End Sub

Private Sub Etiqueta_Change()
    Call objCT.Etiqueta_Change
End Sub

Private Sub Quantidade_Change()
    Call objCT.Quantidade_Change
End Sub

Private Sub Lote_Change()
    Call objCT.Lote_Change
End Sub

Private Sub GridItens_Click()
    Call objCT.GridItens_Click
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridItens_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridItens_EnterCell()
     Call objCT.GridItens_EnterCell
End Sub

Private Sub GridItens_GotFocus()
     Call objCT.GridItens_GotFocus
End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)
     Call objCT.GridItens_KeyPress(KeyAscii)
End Sub

Private Sub GridItens_LeaveCell()
     Call objCT.GridItens_LeaveCell
End Sub

Private Sub GridItens_Validate(Cancel As Boolean)
     Call objCT.GridItens_Validate(Cancel)
End Sub

Private Sub GridItens_Scroll()
     Call objCT.GridItens_Scroll
End Sub

Private Sub Almoxarifado_GotFocus()
     Call objCT.Almoxarifado_GotFocus
End Sub

Private Sub Almoxarifado_KeyPress(KeyAscii As Integer)
     Call objCT.Almoxarifado_KeyPress(KeyAscii)
End Sub

Private Sub Almoxarifado_Validate(Cancel As Boolean)
     Call objCT.Almoxarifado_Validate(Cancel)
End Sub

Private Sub Produto_GotFocus()
     Call objCT.Produto_GotFocus
End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)
     Call objCT.Produto_KeyPress(KeyAscii)
End Sub

Private Sub Produto_Validate(Cancel As Boolean)
     Call objCT.Produto_Validate(Cancel)
End Sub

Private Sub ContaContabil_GotFocus()
     Call objCT.ContaContabil_GotFocus
End Sub

Private Sub ContaContabil_KeyPress(KeyAscii As Integer)
     Call objCT.ContaContabil_KeyPress(KeyAscii)
End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)
     Call objCT.ContaContabil_Validate(Cancel)
End Sub

Private Sub Quantidade_GotFocus()
     Call objCT.Quantidade_GotFocus
End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)
     Call objCT.Quantidade_KeyPress(KeyAscii)
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)
     Call objCT.Quantidade_Validate(Cancel)
End Sub

Private Sub UnidadeMed_GotFocus()
     Call objCT.UnidadeMed_GotFocus
End Sub

Private Sub UnidadeMed_KeyPress(KeyAscii As Integer)
     Call objCT.UnidadeMed_KeyPress(KeyAscii)
End Sub

Private Sub UnidadeMed_Validate(Cancel As Boolean)
     Call objCT.UnidadeMed_Validate(Cancel)
End Sub

Private Sub UnidadeMed_Click()
     Call objCT.UnidadeMed_Click
End Sub

Private Sub CustoUnitario_GotFocus()
     Call objCT.CustoUnitario_GotFocus
End Sub

Private Sub CustoUnitario_KeyPress(KeyAscii As Integer)
     Call objCT.CustoUnitario_KeyPress(KeyAscii)
End Sub

Private Sub CustoUnitario_Validate(Cancel As Boolean)
     Call objCT.CustoUnitario_Validate(Cancel)
End Sub

Private Sub Etiqueta_GotFocus()
     Call objCT.Etiqueta_GotFocus
End Sub

Private Sub Etiqueta_KeyPress(KeyAscii As Integer)
     Call objCT.Etiqueta_KeyPress(KeyAscii)
End Sub

Private Sub Etiqueta_Validate(Cancel As Boolean)
     Call objCT.Etiqueta_Validate(Cancel)
End Sub

Private Sub Tipo_GotFocus()
     Call objCT.Tipo_GotFocus
End Sub

Private Sub Tipo_KeyPress(KeyAscii As Integer)
     Call objCT.Tipo_KeyPress(KeyAscii)
End Sub

Private Sub Tipo_Validate(Cancel As Boolean)
     Call objCT.Tipo_Validate(Cancel)
End Sub





'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_UnLoad(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_UnLoad(Cancel)
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
    Set objCT = New CTInventarioLote
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

'**** fim do trecho a ser copiado *****


Private Sub AlmoxPadraoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(AlmoxPadraoLabel, Source, X, Y)
End Sub

Private Sub AlmoxPadraoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(AlmoxPadraoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
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

Private Sub CodigoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoLabel, Source, X, Y)
End Sub

Private Sub CodigoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoLabel, Button, Shift, X, Y)
End Sub

Private Sub LoteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LoteLabel, Source, X, Y)
End Sub

Private Sub LoteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LoteLabel, Button, Shift, X, Y)
End Sub


'################################################################

Private Sub BotaoLote_Click()
    Call objCT.BotaoLote_Click
End Sub

'######################################################################

'###################################################
'Inserido por Wagner 23/02/2007
Private Sub BotaoProxNum_Click()
    Call objCT.BotaoProxNum_Click
End Sub
'###################################################

