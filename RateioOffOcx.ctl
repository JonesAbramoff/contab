VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl RateioOffOcx 
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9270
   KeyPreview      =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   9270
   Begin VB.Frame FrameRateio 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   3975
      Index           =   1
      Left            =   135
      TabIndex        =   4
      Top             =   1650
      Width           =   6330
      Begin MSMask.MaskEdBox ContaInicio 
         Height          =   225
         Left            =   1245
         TabIndex        =   6
         Top             =   1440
         Width           =   1890
         _ExtentX        =   3334
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
      Begin MSMask.MaskEdBox ContaFim 
         Height          =   225
         Left            =   3165
         TabIndex        =   7
         Top             =   1440
         Width           =   1890
         _ExtentX        =   3334
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
      Begin VB.Frame Frame3 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   720
         Left            =   375
         TabIndex        =   23
         Top             =   3180
         Width           =   4740
         Begin VB.Label ContaOrigemDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   870
            TabIndex        =   25
            Top             =   300
            Width           =   3720
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Left            =   225
            TabIndex        =   26
            Top             =   315
            Width           =   570
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GridContas 
         Height          =   1860
         Left            =   495
         TabIndex        =   8
         Top             =   1065
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSMask.MaskEdBox CclOrigem 
         Height          =   315
         Left            =   1995
         TabIndex        =   5
         Top             =   195
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label LabelCcl 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   540
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   27
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Contas"
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
         Left            =   510
         TabIndex        =   28
         Top             =   825
         Width           =   600
      End
      Begin VB.Label CclOrigemDescricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1980
         TabIndex        =   29
         Top             =   585
         Width           =   4200
      End
   End
   Begin VB.Frame FrameRateio 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3975
      Index           =   2
      Left            =   135
      TabIndex        =   9
      Top             =   1650
      Visible         =   0   'False
      Width           =   6330
      Begin MSMask.MaskEdBox Ccl 
         Height          =   225
         Left            =   2415
         TabIndex        =   13
         Top             =   1770
         Width           =   1410
         _ExtentX        =   2487
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
      Begin MSMask.MaskEdBox Percentual 
         Height          =   225
         Left            =   3525
         TabIndex        =   12
         Top             =   1380
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Format          =   "0%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Conta 
         Height          =   225
         Left            =   1485
         TabIndex        =   11
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Frame Frame2 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   450
         TabIndex        =   24
         Top             =   2880
         Width           =   5790
         Begin VB.Label CclLabel 
            AutoSize        =   -1  'True
            Caption         =   "CCusto:"
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
            Left            =   120
            TabIndex        =   30
            Top             =   660
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label CclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   870
            TabIndex        =   31
            Top             =   645
            Visible         =   0   'False
            Width           =   4785
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Left            =   225
            TabIndex        =   32
            Top             =   315
            Width           =   570
         End
         Begin VB.Label ContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   870
            TabIndex        =   33
            Top             =   300
            Width           =   4770
         End
      End
      Begin MSMask.MaskEdBox ContaCredito 
         Height          =   315
         Left            =   1800
         TabIndex        =   10
         Top             =   120
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSFlexGridLib.MSFlexGrid GridLancamentos 
         Height          =   1860
         Left            =   480
         TabIndex        =   14
         Top             =   690
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label ContaCreditoDescricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3375
         TabIndex        =   34
         Top             =   120
         Width           =   2865
      End
      Begin VB.Label LabelContaCredito 
         AutoSize        =   -1  'True
         Caption         =   "Conta Crédito:"
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
         Left            =   510
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   35
         Top             =   165
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rateios"
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
         Left            =   525
         TabIndex        =   36
         Top             =   495
         Width           =   660
      End
      Begin VB.Label TotalPercentual 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2400
         TabIndex        =   37
         Top             =   2535
         Width           =   1155
      End
      Begin VB.Label LabelTotais 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Height          =   195
         Left            =   1665
         TabIndex        =   38
         Top             =   2535
         Width           =   600
      End
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
      Left            =   6780
      TabIndex        =   45
      Top             =   3630
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
      Left            =   6780
      TabIndex        =   44
      Top             =   2625
      Width           =   1605
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1635
      Picture         =   "RateioOffOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   165
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7020
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RateioOffOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RateioOffOcx.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RateioOffOcx.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RateioOffOcx.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox TipoRateio 
      Height          =   315
      ItemData        =   "RateioOffOcx.ctx":0A7E
      Left            =   945
      List            =   "RateioOffOcx.ctx":0A88
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   750
      Width           =   2430
   End
   Begin MSComctlLib.TreeView TvwContas 
      Height          =   4080
      Left            =   6000
      TabIndex        =   15
      Top             =   4545
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   7197
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
   Begin MSComctlLib.TreeView TvwCcls 
      Height          =   4080
      Left            =   5880
      TabIndex        =   16
      Top             =   4635
      Visible         =   0   'False
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   7197
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
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   945
      TabIndex        =   0
      Top             =   150
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   556
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
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   3705
      TabIndex        =   2
      Top             =   120
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4365
      Left            =   120
      TabIndex        =   22
      Top             =   1305
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7699
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Origem"
            Object.ToolTipText     =   "Centro de Custo/Contas que serão a origem do valor a ser rateado"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Destino"
            Object.ToolTipText     =   "Contas/Centros de Custo que serão o destino do valor do rateio"
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
   Begin VB.Label LabelContas 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6060
      TabIndex        =   39
      Top             =   4035
      Width           =   2730
   End
   Begin VB.Label LabelCcls 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6045
      TabIndex        =   40
      Top             =   3705
      Width           =   2160
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   435
      TabIndex        =   41
      Top             =   795
      Width           =   450
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
      Left            =   210
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   42
      Top             =   165
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2685
      TabIndex        =   43
      Top             =   180
      Width           =   945
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   -285
      X2              =   9540
      Y1              =   1215
      Y2              =   1215
   End
End
Attribute VB_Name = "RateioOffOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()
Dim iFrameAtual As Integer

Dim iGrid_Conta_Col As Integer
Dim iGrid_Ccl_Col As Integer
Dim iGrid_Percentual_Col As Integer

Dim iGrid_ContaInicio_Col As Integer
Dim iGrid_ContaFinal_Col As Integer

Dim objGrid1 As AdmGrid
Dim objGridContas As AdmGrid
Dim iAlterado As Integer

Private WithEvents objEventoRateioOff As AdmEvento
Attribute objEventoRateioOff.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoConta As AdmEvento
Attribute objEventoConta.VB_VarHelpID = -1

Const TESTA_LINHA_ATUAL = 1
Const NAO_TESTA_LINHA_ATUAL = 0
Const TAB_ORIGEM = 1
Const TAB_DESTINO = 2


Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Mostra número do proximo rateio disponível
    lErro = CF("RateioOff_Automatico", lCodigo)
    If lErro <> SUCESSO Then Error 57517

    Codigo.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57517
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166110)
    
    End Select

    Exit Sub

End Sub

Private Sub Ccl_GotFocus()
    
Dim lErro As Long
Dim sCcl As String

On Error GoTo Erro_Ccl_GotFocus

'    TvwContas.Visible = False
'    LabelContas.Visible = False
'    TvwCcls.Visible = True
'    LabelCcls.Visible = True
    BotaoCcl.Tag = "Ccl"
    
    CclDescricao.Caption = ""

    Call Grid_Campo_Recebe_Foco(objGrid1)
      
    sCcl = GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col)
    
    If Len(sCcl) > 0 Then

        lErro = Ccl_Exibe_Descricao(sCcl)
        If lErro <> SUCESSO Then Error 20566
        
    End If
    
    Exit Sub
    
Erro_Ccl_GotFocus:

    Select Case Err
    
        Case 20566
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166111)
    
    End Select
    
    Exit Sub
    
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
    If lErro <> SUCESSO Then Error 20567
    
    objCcl.sCcl = sCclFormatada

    lErro = CF("Ccl_Le", objCcl)
    If lErro <> SUCESSO And lErro <> 5599 Then Error 20568
    
    If lErro = 5599 Then Error 20569
    
    CclDescricao.Caption = objCcl.sDescCcl
    
    Ccl_Exibe_Descricao = SUCESSO
    
    Exit Function

Erro_Ccl_Exibe_Descricao:

    Ccl_Exibe_Descricao = Err
    
    Select Case Err
    
        Case 20567, 20568
            CclDescricao = ""
            
        Case 20569
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", Err, objCcl.sCcl)
            CclDescricao = ""
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166112)
            
    End Select
        
    Exit Function

End Function

Private Sub CclOrigem_GotFocus()
    
    BotaoCcl.Tag = "CclOrigem"
'    TvwContas.Visible = False
'    LabelContas.Visible = False
'    TvwCcls.Visible = True
'    LabelCcls.Visible = True
    
End Sub

Private Sub CclOrigem_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CclOrigem_Validate

    'critica o formato do ccl, sua presença no BD e capacidade de receber lançamentos
    lErro = CF("Ccl_Critica", CclOrigem.Text, sCclFormatada, objCcl)
    If lErro <> SUCESSO And lErro <> 5703 Then Error 18197
                
    'se o centro de custo/lucro não estiver cadastrado
    If lErro = 5703 Then Error 18198
    
    CclOrigemDescricao.Caption = objCcl.sDescCcl
    
    Exit Sub
    
Erro_CclOrigem_Validate:

    Cancel = True

    If Not (Parent Is GL_objMDIForm.ActiveForm) Then
        Me.Show
    End If

    Select Case Err
    
        Case 18197
            CclOrigem.SetFocus
                    
        Case 18198
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", CclOrigem.Text)

            If vbMsgRes = vbYes Then
            
                objCcl.sCcl = sCclFormatada
                
                Call Chama_Tela("CclTela", objCcl)
       
            End If
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166113)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub ContaCredito_GotFocus()

    BotaoConta.Tag = "ContaCredito"
'    TvwContas.Visible = True
'    LabelContas.Visible = True
'    TvwCcls.Visible = False
'    LabelCcls.Visible = False
    
End Sub

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Form_Load

    iFrameAtual = TAB_ORIGEM
  
    Set objGrid1 = New AdmGrid
    Set objGridContas = New AdmGrid
    
    Set objEventoRateioOff = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoConta = New AdmEvento
    
    TvwContas.Visible = False
    LabelContas.Visible = False
    TvwCcls.Visible = False
    LabelCcls.Visible = False
    
    'tela em questão
    Set objGrid1.objForm = Me
  
    If giSetupUsoCcl <> CCL_USA_EXTRACONTABIL And giSetupUsoCcl <> CCL_USA_CONTABIL Then Error 36773
    
'    'Inicializa a Lista de Centros de Custo
'    lErro = Carga_Arvore_Ccl(TvwCcls.Nodes)
'    If lErro <> SUCESSO Then Error 24009
        
    TipoRateio.ListIndex = 0
    
'    'Inicializa a Lista de Plano de Contas
'    lErro = CF("Carga_Arvore_Conta", TvwContas.Nodes)
'    If lErro <> SUCESSO Then Error 11247
    
    lErro = Inicializa_Grid_Lancamentos(objGrid1)
    If lErro <> SUCESSO Then Error 11246
    
    'Inicializa Grid de Contas
    lErro = Inicializa_Grid_Contas(objGridContas)
    If lErro <> SUCESSO Then Error 55728
    
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        CclLabel.Visible = True
        CclDescricao.Visible = True
    End If
    
    TotalPercentual.Caption = Format(0, "Percent")
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case 11246, 11247, 24009, 55728
        
        Case 36773
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_USADO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166114)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim objRateioOff As New ClassRateioOff
Dim colSelecao As Collection

    If Len(Codigo.Text) = 0 Then
        objRateioOff.lCodigo = 0
    Else
        objRateioOff.lCodigo = CInt(Codigo.ClipText)
    End If

    objRateioOff.lSeq = 0

    Call Chama_Tela("RateioOffLista", colSelecao, objRateioOff, objEventoRateioOff)
   
End Sub

Private Sub LabelContaCredito_Click()
    BotaoConta.Tag = "ContaCredito"
    Call BotaoConta_Click
End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        FrameRateio(Opcao.SelectedItem.Index).Visible = True
        FrameRateio(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

        If iFrameAtual = TAB_ORIGEM Then
            BotaoConta.Tag = "ContaInicio"
            BotaoCcl.Tag = "CclOrigem"
        Else
            BotaoConta.Tag = "ContaCredito"
            BotaoCcl.Tag = "Ccl"
        End If

    End If

End Sub

Private Sub Percentual_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_TvwContas_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then

        'move os dados do plano de contas do banco de dados para a arvore colNodes.
        lErro = CF("Carga_Arvore_Conta1", objNode, TvwContas.Nodes)
        If lErro <> SUCESSO Then Error 36772

    End If
    
    Exit Sub
    
Erro_TvwContas_Expand:

    Select Case Err
    
        Case 36772
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166116)
        
    End Select
        
    Exit Sub
    
End Sub

Private Function Inicializa_Grid_Lancamentos(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_Lancamentos
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Conta")
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then objGridInt.colColuna.Add ("CCusto")
    objGridInt.colColuna.Add ("Percentual")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (Conta.Name)
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (Percentual.Name)
    
    'indica onde estão situadas as colunas do grid
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        iGrid_Conta_Col = 1
        iGrid_Ccl_Col = 2
        iGrid_Percentual_Col = 3
    Else
        iGrid_Conta_Col = 1
        '999 indica que não está sendo usado
        iGrid_Ccl_Col = 999
        iGrid_Percentual_Col = 2
        Ccl.Visible = False
    End If
    
    lErro = Inicializa_Mascaras()
    If lErro <> SUCESSO Then Error 11251
    
    objGridInt.objGrid = GridLancamentos
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 501
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 6
        
    GridLancamentos.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGridInt)

    'Posiciona os painéis totalizadores
    TotalPercentual.Top = GridLancamentos.Top + GridLancamentos.Height
    TotalPercentual.Left = GridLancamentos.Left
    For iIndice = 0 To iGrid_Percentual_Col - 1
        TotalPercentual.Left = TotalPercentual.Left + GridLancamentos.ColWidth(iIndice) + GridLancamentos.GridLineWidth + 20
    Next

    TotalPercentual.Width = GridLancamentos.ColWidth(iGrid_Percentual_Col)
    
    LabelTotais.Top = TotalPercentual.Top + (TotalPercentual.Height - LabelTotais.Height) / 2
    LabelTotais.Left = TotalPercentual.Left - LabelTotais.Width

    Inicializa_Grid_Lancamentos = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_Lancamentos:

    Inicializa_Grid_Lancamentos = Err
    
    Select Case Err
    
        Case 11251
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166117)
        
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
    If lErro <> SUCESSO Then Error 11252
    
    Conta.Mask = sMascaraConta
    ContaCredito.Mask = sMascaraConta
    ContaInicio.Mask = sMascaraConta
    ContaFim.Mask = sMascaraConta
   
    'Se usa centro de custo/lucro ==> inicializa mascara de centro de custo/lucro
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Or giSetupUsoCcl = CCL_USA_CONTABIL Then
    
        sMascaraCcl = String(STRING_CCL, 0)

        'le a mascara dos centros de custo/lucro
        lErro = MascaraCcl(sMascaraCcl)
        If lErro <> SUCESSO Then Error 11135

        CclOrigem.Mask = sMascaraCcl
        
    End If
    
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then Ccl.Mask = sMascaraCcl

    Inicializa_Mascaras = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascaras:

    Inicializa_Mascaras = Err
    
    Select Case Err
    
        Case 11252, 11135
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166118)
        
    End Select

    Exit Function
    
End Function

Function Trata_Parametros(Optional objRateioOff As ClassRateioOff) As Long

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um Rateio passado como parametro, exibir seus dados
    If Not (objRateioOff Is Nothing) Then
    
        lErro = Traz_Doc_Tela(objRateioOff)
        If lErro <> SUCESSO Then Error 11253
    
    Else
    
        iAlterado = 0
    
    End If
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case 11253
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166119)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Private Function Traz_Doc_Tela(objRateioOff As ClassRateioOff) As Long
'traz os dados do rateio do banco de dados para a tela

Dim lErro As Long
Dim colRateioOff As New Collection
Dim objRateioOff1 As ClassRateioOff
Dim sContaMascarada As String
Dim sCclMascarado As String
Dim dTotal As Double
Dim iIndice As Integer
Dim objCcl As New ClassCcl
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_Traz_Doc_Tela

    Call Limpa_Tela_RateioOff
        
    'Coloca em colRateioOff Todos os dados do rateio passado em objRateioOff
    lErro = CF("RateioOff_Le_Doc", objRateioOff, colRateioOff)
    If lErro <> SUCESSO And lErro <> 11345 Then Error 11254
        
    'se não está cadastrado
    If lErro = 11345 Then Error 18191
    
    Codigo.Text = CStr(objRateioOff.lCodigo)
        
    'mascara o centro de custo
    sCclMascarado = String(STRING_CCL, 0)

    lErro = Mascara_MascararCcl(objRateioOff.sCclOrigem, sCclMascarado)
    If lErro <> SUCESSO Then Error 20576
    
    CclOrigem.PromptInclude = False
    CclOrigem.Text = sCclMascarado
    CclOrigem.PromptInclude = True
        
    objCcl.sCcl = objRateioOff.sCclOrigem
        
    lErro = CF("Ccl_Le", objCcl)
    If lErro <> SUCESSO And lErro <> 5599 Then Error 55816
    
    If lErro = 5599 Then Error 55817
    
    CclOrigemDescricao.Caption = objCcl.sDescCcl
        
    Descricao.Text = objRateioOff.sDescricao
        
    For Each objRateioOff1 In colRateioOff
     
        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)
         
        lErro = Mascara_RetornaContaEnxuta(objRateioOff1.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 11257
         
        'move as conta envolvidas no rateio para o Grid
        Conta.PromptInclude = False
        Conta.Text = sContaMascarada
        Conta.PromptInclude = True
         
        'coloca a conta na tela
        GridLancamentos.TextMatrix(objRateioOff1.lSeq, iGrid_Conta_Col) = Conta.Text
                  
        'coloca o percentual na tela
        GridLancamentos.TextMatrix(objRateioOff1.lSeq, iGrid_Percentual_Col) = Format(objRateioOff1.dPercentual, "Percent")
         
        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        
            'mascara o centro de custo
            sCclMascarado = String(STRING_CCL, 0)
               
            If objRateioOff1.sCcl <> "" Then
            
                lErro = Mascara_MascararCcl(objRateioOff1.sCcl, sCclMascarado)
                If lErro <> SUCESSO Then Error 20575
        
            Else
              sCclMascarado = ""
                   
            End If
                
            'coloca o centro de custo na tela
            GridLancamentos.TextMatrix(objRateioOff1.lSeq, iGrid_Ccl_Col) = sCclMascarado
            
        End If
         
        objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
            
    Next
     
    'Coloca o Total na tela
    dTotal = GridColuna_Soma(iGrid_Percentual_Col)
    TotalPercentual.Caption = Format(dTotal, "Percent")
        
    sContaMascarada = String(STRING_CONTA, 0)
          
    lErro = Mascara_RetornaContaEnxuta(objRateioOff.sContaCre, sContaMascarada)
    If lErro <> SUCESSO Then Error 11258
        
    'Coloca a conta credito na tela
    ContaCredito.PromptInclude = False
    ContaCredito.Text = sContaMascarada
    ContaCredito.PromptInclude = True

    objCcl.sCcl = objRateioOff.sCclOrigem
        
    lErro = CF("PlanoConta_Le_Conta1", objRateioOff.sContaCre, objPlanoConta)
    If lErro <> SUCESSO And lErro <> 6030 Then Error 55818
    
    If lErro = 6030 Then Error 55819

    ContaCreditoDescricao.Caption = objPlanoConta.sDescConta

    'Coloca na Tela o tipo do Rateio
    For iIndice = 0 To TipoRateio.ListCount - 1
    
        If TipoRateio.ItemData(iIndice) = objRateioOff.iTipo Then
            TipoRateio.ListIndex = iIndice
            Exit For
        End If
        
    Next
                
    'traz as contas origem do rateio
    lErro = Traz_Doc_Tela1(objRateioOff)
    If lErro <> SUCESSO Then Error 55807
                
    iAlterado = 0
    
    Traz_Doc_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Doc_Tela:

    Traz_Doc_Tela = Err

    Select Case Err
        
        Case 11254, 55807, 55816, 55818
        
        Case 11257
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objRateioOff1.sConta)
            
        Case 11258
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objRateioOff.sContaCre)
        
        Case 18191
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RATEIOOFF_NAO_CADASTRADO", Err)
        
        Case 20575, 20576
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objRateioOff1.sCcl)
        
        Case 55817
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", Err, objCcl.sCcl)
        
        Case 55819
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, objRateioOff1.sContaCre)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166120)
        
    End Select
    
    iAlterado = 0
    
    Exit Function
        
End Function

Private Function Traz_Doc_Tela1(objRateioOff As ClassRateioOff) As Long
'traz as contas origem do rateio

Dim lErro As Long
Dim colContas As New Collection
Dim objRateioOffContas As ClassRateioOffContas
Dim sContaMascarada As String
Dim iIndice As Integer

On Error GoTo Erro_Traz_Doc_Tela1

    'Coloca em colRateioOff Todos os dados do rateio passado em objRateioOff
    lErro = CF("RateioOffContas_Le_Doc", objRateioOff, colContas)
    If lErro <> SUCESSO Then Error 55810
        
    'se não está cadastrado
    If lErro = SUCESSO Then
    
        For Each objRateioOffContas In colContas
         
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
             
            lErro = Mascara_RetornaContaEnxuta(objRateioOffContas.sContaInicio, sContaMascarada)
            If lErro <> SUCESSO Then Error 55808
             
            'move as conta envolvidas no rateio para o Grid
            ContaInicio.PromptInclude = False
            ContaInicio.Text = sContaMascarada
            ContaInicio.PromptInclude = True
             
            'coloca a conta na tela
            GridContas.TextMatrix(objRateioOffContas.iItem, iGrid_ContaInicio_Col) = ContaInicio.Text
                      
            lErro = Mascara_RetornaContaEnxuta(objRateioOffContas.sContaFim, sContaMascarada)
            If lErro <> SUCESSO Then Error 55809
             
            'move as conta envolvidas no rateio para o Grid
            ContaFim.PromptInclude = False
            ContaFim.Text = sContaMascarada
            ContaFim.PromptInclude = True
             
            'coloca a conta na tela
            GridContas.TextMatrix(objRateioOffContas.iItem, iGrid_ContaFinal_Col) = ContaFim.Text
             
            objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
                
        Next
     
    End If
    
    Traz_Doc_Tela1 = SUCESSO
    
    Exit Function
    
Erro_Traz_Doc_Tela1:

    Traz_Doc_Tela1 = Err

    Select Case Err
        
        Case 55808
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objRateioOffContas.sContaInicio)
            
        Case 55809
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objRateioOffContas.sContaFim)
        
        Case 55810
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166121)
        
    End Select
    
    iAlterado = 0
    
    Exit Function
        
End Function

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Conta_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CclOrigem_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ContaCredito_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ContaCredito_Validate(Cancel As Boolean)
    
Dim iLinha As Integer
Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada  As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim sConta As String
Dim sContaFormatadaGrid As String
Dim iContaPreenchida As Integer

On Error GoTo Erro_ContaCredito_Validate
        
    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica", ContaCredito.Text, ContaCredito.ClipText, objPlanoConta)
    If lErro <> SUCESSO And lErro <> 44033 And lErro <> 44037 Then Error 19367
    
    'se é uma conta simples, coloca a conta normal no lugar da conta simples
    If lErro = SUCESSO Then
    
        sContaFormatada = objPlanoConta.sConta
        
        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)
        
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 19368
        
        ContaCredito.PromptInclude = False
        ContaCredito.Text = sContaMascarada
        ContaCredito.PromptInclude = True
        
    'se não encontrou a conta simples
    ElseIf lErro = 44033 Or lErro = 44037 Then
    
        'testa a conta no seu formato normal
        'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaCredito.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 11870

        'conta não cadastrada
        If lErro = 5700 Then Error 11871
        
    End If
    
    'verifica se a conta passada como parametro coincide com as contas em lançamentos
'    lErro = Testa_Conta_Grid_Lancamentos(ContaCredito.Text, TESTA_LINHA_ATUAL)
'    If lErro <> SUCESSO Then Error 55770
        
    'verifica se a conta passada como parametro coincide com as contas em lançamentos
'    lErro = Testa_Conta_Grid_Contas(ContaCredito.Text, TESTA_LINHA_ATUAL)
'    If lErro <> SUCESSO Then Error 55771
        
    ContaCreditoDescricao.Caption = objPlanoConta.sDescConta
         
    Exit Sub
    
Erro_ContaCredito_Validate:

    Cancel = True

    If Not (Parent Is GL_objMDIForm.ActiveForm) Then
        Me.Show
    End If

    Select Case Err
        
        Case 11870, 19367, 55770, 55771
            ContaCredito.SetFocus
            
        Case 11871
             vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaCredito.Text)
            
            If vbMsgRes = vbYes Then
                
                objPlanoConta.sConta = sContaFormatada
                
                'Usuário quer criar esta conta
                Call Chama_Tela("PlanoConta", objPlanoConta)
            
            Else
                ContaCredito.SetFocus
            End If
        
        Case 18192
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_JA_UTILIZADA", Err, ContaCredito.Text)
            ContaCredito.SetFocus
            
        Case 19368
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            ContaCredito.SetFocus
            
        Case 55674
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166122)
    
    End Select
    
    Exit Sub
        
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
    
    Set objEventoRateioOff = Nothing
    Set objEventoCcl = Nothing
    Set objEventoConta = Nothing
    
    Set objGrid1 = Nothing
    
    Set objGridContas = Nothing

End Sub

Private Sub GridLancamentos_LeaveCell()
    Call Saida_Celula(objGrid1)
End Sub

Private Sub GridLancamentos_EnterCell()
    Call Grid_Entrada_Celula(objGrid1, iAlterado)
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

Private Sub GridLancamentos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid1)
    TotalPercentual.Caption = Format(GridColuna_Soma(iGrid_Percentual_Col), "Percent")
    
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

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then
    
        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o Grid de Rateio Destino
            Case GridLancamentos.Name

                lErro = Saida_Celula_Lancamentos(objGridInt)
                If lErro <> SUCESSO Then Error 55729

            'Se for o Grid de Contas Origem
            Case GridContas.Name

                lErro = Saida_Celula_Contas(objGridInt)
                If lErro <> SUCESSO Then Error 55730

        End Select
    
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 11262
        
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula:

    Saida_Celula = Err
    
    Select Case Err
            
        Case 55729, 55730
    
        Case 11262
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166123)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Lancamentos(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Lancamentos
    
    Select Case GridLancamentos.Col

        Case iGrid_Conta_Col
        
            lErro = Saida_Celula_Conta(objGridInt)
            If lErro <> SUCESSO Then Error 11260
            
        Case iGrid_Ccl_Col
        
            lErro = Saida_Celula_Ccl(objGridInt)
            If lErro <> SUCESSO Then Error 20558
            
        Case iGrid_Percentual_Col
        
            lErro = Saida_Celula_Percentual(objGridInt)
            If lErro <> SUCESSO Then Error 11261
            
    End Select

    Saida_Celula_Lancamentos = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Lancamentos:

    Saida_Celula_Lancamentos = Err
    
    Select Case Err
            
        Case 11260, 11261, 20558
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166124)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Conta(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim sContaFormatada As String
Dim sContaMascarada As String
Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim objContaCcl As New ClassContaCcl

On Error GoTo Erro_Saida_Celula_Conta

    Set objGridInt.objControle = Conta
    
    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica", Conta.Text, Conta.ClipText, objPlanoConta)
    If lErro <> SUCESSO And lErro <> 44033 And lErro <> 44037 Then Error 19365
    
    'se é uma conta simples, coloca a conta normal no lugar da conta simples
    If lErro = SUCESSO Then
    
        sContaFormatada = objPlanoConta.sConta
        
        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)
        
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 19366
        
        Conta.PromptInclude = False
        Conta.Text = sContaMascarada
        Conta.PromptInclude = True
        
    'se não encontrou a conta simples
    ElseIf lErro = 44033 Or lErro = 44037 Then

        'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", Conta.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 11263
                    
        'conta não cadastrada
        If lErro = 5700 Then Error 11264
                
    End If
    
    'Verifica se a conta foi preenchida
    If Len(Conta.ClipText) > 0 Then
    
'        'verifica se a conta passada como parametro coincide com a conta credito
'        lErro = Testa_Conta_Credito(Conta.Text)
'        If lErro <> SUCESSO Then Error 55761
'
        'verifica se a conta passada como parametro coincide com as contas em lançamentos
'        lErro = Testa_Conta_Grid_Contas(Conta.Text, TESTA_LINHA_ATUAL)
'        If lErro <> SUCESSO Then Error 55763
        
        'verifica se a conta passada como parametro tem associacao com o centro de custo em questao
        lErro = Testa_Assoc_ContaCcl(sContaFormatada, objContaCcl)
        If lErro <> SUCESSO And lErro <> 20557 Then Error 55757
        
        'se está faltando a associacao da conta com o centro de custo
        If lErro = 20557 Then Error 55758
        
        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

        'Coloca a descricao daconta na tela
        ContaDescricao.Caption = objPlanoConta.sDescConta

    End If
                        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 11270

    Saida_Celula_Conta = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Conta:

    Saida_Celula_Conta = Err
    
    Select Case Err
            
        Case 11263, 11270, 19365, 55757, 55761, 55763
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 11264
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", Conta.Text)

            If vbMsgRes = vbYes Then
            
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("PlanoConta", objPlanoConta)
                
            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If
                      
        Case 19366
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 55758
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTACCL_INEXISTENTE", Conta.Text, GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Ccl_Col))

            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
            
                Call Chama_Tela("ContaCcl", objContaCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166125)
    
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
    If lErro <> SUCESSO And lErro <> 5703 Then Error 20559
                
    'se o centro de custo/lucro não estiver cadastrado
    If lErro = 5703 Then Error 20560
                
    'se o centro de custo foi preenchido
    If Len(Ccl.ClipText) > 0 Then
    
        'se a conta foi informada
        If Len(GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Conta_Col)) > 0 Then
    
            'verificar se a associação da conta com o centro de custo em questão está cadastrada
            sConta = GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Conta_Col)
        
            lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
            If lErro <> SUCESSO Then Error 20561
        
            objContaCcl.sConta = sContaFormatada
            objContaCcl.sCcl = sCclFormatada
        
            lErro = CF("ContaCcl_Le", objContaCcl)
            If lErro <> SUCESSO And lErro <> 5871 Then Error 20562
        
            'associação Conta x Centro de Custo/Lucro não cadastrada
            If lErro = 5871 Then Error 20563
        
        End If
                        
        CclDescricao.Caption = objCcl.sDescCcl
             
'        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.ilinhasExistentes Then
'            objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
'        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 20564

    Saida_Celula_Ccl = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = Err
    
    Select Case Err
    
        Case 20559, 20561, 20562, 20564
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 20560
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)

            If vbMsgRes = vbYes Then
            
                objCcl.sCcl = sCclFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("CclTela", objCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If
            
        Case 20563
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTACCL_INEXISTENTE", sConta, Ccl.Text)

            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
            
                Call Chama_Tela("ContaCcl", objContaCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166126)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Percentual(objGridInt As AdmGrid) As Long
'faz a critica da celula de percentual do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dColunaSoma As Double
Dim dValor As Double

On Error GoTo Erro_Saida_Celula_Percentual

    Set objGridInt.objControle = Percentual
    
    'Verifica se o percentual foi preenchido
    If Len(Percentual.ClipText) > 0 Then

        lErro = Porcentagem_Critica(Percentual.Text)
        If lErro <> SUCESSO Then Error 11271

        dValor = CDbl(Percentual.Text)

        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
     End If
                  
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 11272
                        
    dColunaSoma = GridColuna_Soma(iGrid_Percentual_Col)
    TotalPercentual.Caption = Format(dColunaSoma, "Percent")
      
    Saida_Celula_Percentual = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Percentual:

    Saida_Celula_Percentual = Err
    
    Select Case Err
    
        Case 11271
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 11272
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166127)
        
    End Select

    Exit Function

End Function

Function GridColuna_Soma(iColuna As Integer) As Double
'Rotina Responsavel por somar os valores dos percentuais
    
Dim dAcumulador As Double
Dim iLinha As Integer
    
    dAcumulador = 0
    
    'Para cadalinha existente ele acumula o valor do percentual
    For iLinha = 1 To objGrid1.iLinhasExistentes
        If Len(GridLancamentos.TextMatrix(iLinha, iColuna)) > 0 Then
            dAcumulador = dAcumulador + CDbl(Format(GridLancamentos.TextMatrix(iLinha, iColuna), "General Number"))
        End If
    Next
    
    GridColuna_Soma = dAcumulador

End Function

Private Sub LabelCcl_Click()
    BotaoCcl.Tag = "CclOrigem"
    Call BotaoCcl_Click
End Sub

Private Sub objEventoRateioOff_evSelecao(obj1 As Object)
    
Dim objRateioOff As ClassRateioOff
Dim lErro As Long
    
On Error GoTo Erro_objEventoRateioOff_evSelecao
    
    Set objRateioOff = obj1
    
    lErro = Traz_Doc_Tela(objRateioOff)
    If lErro <> SUCESSO Then Error 11120

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
                
    iAlterado = 0
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoRateioOff_evSelecao:

    Select Case Err
    
        Case 11120
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166129)
            
    End Select
        
    Exit Sub
        
End Sub

Private Sub Percentual_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Percentual_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid1)
End Sub

Private Sub Percentual_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Percentual
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub TipoRateio_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
    
Dim sConta As String
Dim sCaracterInicial As String
Dim lPosicaoSeparador As Long
Dim lErro As Long
Dim sContaEnxuta As String
Dim vbMsgRes As VbMsgBoxResult
Dim objContaCcl As New ClassContaCcl
Dim sContaMascarada As String

On Error GoTo Erro_TvwContas_NodeClick

    ContaDescricao.Caption = ""
    
    sCaracterInicial = Left(Node.Key, 1)

    If sCaracterInicial = "A" Then

        sConta = Right(Node.Key, Len(Node.Key) - 1)
        
        lPosicaoSeparador = InStr(Node.Text, SEPARADOR)
        sContaMascarada = Mid(Node.Text, 1, lPosicaoSeparador - 1)

        sContaEnxuta = String(STRING_CONTA, 0)
        
        'volta mascarado apenas os caracteres preenchidos
        lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 11273
                
        If TvwContas.Tag = "Conta" Then
                
            If GridLancamentos.Col = iGrid_Conta_Col Then
            
                If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.iLinhasExistentes Then
                    objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
                End If
            
'                'verifica se a conta passada como parametro coincide com a conta credito
'                lErro = Testa_Conta_Credito(sContaMascarada)
'                If lErro <> SUCESSO Then Error 55742
'
'                'verifica se a conta passada como parametro coincide com as contas em lançamentos
'                lErro = Testa_Conta_Grid_Contas(sContaMascarada, TESTA_LINHA_ATUAL)
'                If lErro <> SUCESSO Then Error 55747

                'verifica se a conta passada como parametro tem associacao com o centro de custo em questao
                lErro = Testa_Assoc_ContaCcl(sContaMascarada, objContaCcl)
                If lErro <> SUCESSO And lErro <> 20557 Then Error 55759
                
                'se está faltando a associacao da conta com o centro de custo
                If lErro = 20557 Then Error 55760

                Conta.PromptInclude = False
                Conta.Text = sContaEnxuta
                Conta.PromptInclude = True

                GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Conta_Col) = Conta.Text

                'Preenche a Descricao da Conta
                lPosicaoSeparador = InStr(Node.Text, SEPARADOR)
                ContaDescricao.Caption = Mid(Node.Text, lPosicaoSeparador + 1)

            End If
            
        ElseIf TvwContas.Tag = "ContaCredito" Then

'            'verifica se a conta passada como parametro coincide com as contas em lançamentos
'            lErro = Testa_Conta_Grid_Lancamentos(sContaMascarada, TESTA_LINHA_ATUAL)
'            If lErro <> SUCESSO Then Error 55739
'
'            'verifica se a conta passada como parametro coincide com as contas em lançamentos
'            lErro = Testa_Conta_Grid_Contas(sContaMascarada, TESTA_LINHA_ATUAL)
'            If lErro <> SUCESSO Then Error 55748
            
            ContaCredito.PromptInclude = False
            ContaCredito.Text = sContaEnxuta
            ContaCredito.PromptInclude = True
        
            'Preenche a Descricao da Conta
            lPosicaoSeparador = InStr(Node.Text, SEPARADOR)
            ContaCreditoDescricao.Caption = Mid(Node.Text, lPosicaoSeparador + 1)
        
        ElseIf TvwContas.Tag = "ContaInicio" Then
                
            If GridContas.Col = iGrid_ContaInicio_Col Then
            
                If objGridContas.objGrid.Row - objGridContas.objGrid.FixedRows = objGridContas.iLinhasExistentes Then
                    objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
                End If
            
'                'verifica se a conta passada como parametro coincide com a conta credito
'                lErro = Testa_Conta_Credito(sContaMascarada)
'                If lErro <> SUCESSO Then Error 55743
'
'                'verifica se a conta passada como parametro coincide com as contas em lançamentos
'                lErro = Testa_Conta_Grid_Lancamentos(sContaMascarada, TESTA_LINHA_ATUAL)
'                If lErro <> SUCESSO Then Error 55740
'
'                'verifica se a conta passada como parametro coincide com as contas em lançamentos
'                lErro = Testa_Conta_Grid_Contas(sContaMascarada, NAO_TESTA_LINHA_ATUAL)
'                If lErro <> SUCESSO Then Error 55749

                ContaInicio.PromptInclude = False
                ContaInicio.Text = sContaEnxuta
                ContaInicio.PromptInclude = True

                GridContas.TextMatrix(GridContas.Row, iGrid_ContaInicio_Col) = ContaInicio.Text

                'Preenche a Descricao da Conta
                lPosicaoSeparador = InStr(Node.Text, SEPARADOR)
                ContaOrigemDescricao.Caption = Mid(Node.Text, lPosicaoSeparador + 1)
                
            End If

        ElseIf TvwContas.Tag = "ContaFim" Then
                
            If GridContas.Col = iGrid_ContaFinal_Col Then
            
                If objGridContas.objGrid.Row - objGridContas.objGrid.FixedRows = objGridContas.iLinhasExistentes Then
                    objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
                End If
            
'                'verifica se a conta passada como parametro coincide com a conta credito
'                lErro = Testa_Conta_Credito(sContaMascarada)
'                If lErro <> SUCESSO Then Error 55750
'
'                'verifica se a conta passada como parametro coincide com as contas em lançamentos
'                lErro = Testa_Conta_Grid_Lancamentos(sContaMascarada, TESTA_LINHA_ATUAL)
'                If lErro <> SUCESSO Then Error 55751
'
'                'verifica se a conta passada como parametro coincide com as contas em lançamentos
'                lErro = Testa_Conta_Grid_Contas(sContaMascarada, NAO_TESTA_LINHA_ATUAL)
'                If lErro <> SUCESSO Then Error 55752

                ContaFim.PromptInclude = False
                ContaFim.Text = sContaEnxuta
                ContaFim.PromptInclude = True

                GridContas.TextMatrix(GridContas.Row, iGrid_ContaFinal_Col) = ContaFim.Text

                'Preenche a Descricao da Conta
                lPosicaoSeparador = InStr(Node.Text, SEPARADOR)
                ContaOrigemDescricao.Caption = Mid(Node.Text, lPosicaoSeparador + 1)

            End If
        
        End If
    
    End If
        
    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case Err
                
        Case 11273
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)
        
        Case 55739, 55740, 55742, 55743, 55747, 55748, 55749, 55750, 55751, 55752, 55759
        
        Case 55760
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTACCL_INEXISTENTE", sConta, GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Ccl_Col))

            If vbMsgRes = vbYes Then
                Call Chama_Tela("ContaCcl", objContaCcl)
            End If
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166130)
            
    End Select
        
    Exit Sub
    
End Sub

Private Function Testa_Conta_Grid_Contas(sConta As String, iTestaLinhaAtual As Integer) As Long
'verifica se a conta passada como parametro coincide com as contas em lançamentos

Dim lErro As Long
Dim iLinha As Integer
Dim sContaFormatada As String
Dim sContaInicioFormatadaGrid As String
Dim sContaFinalFormatadaGrid As String
Dim iContaPreenchida As Integer
Dim iContaPreenchida1 As Integer
Dim iLinhaAtual As Integer

On Error GoTo Erro_Testa_Conta_Grid_Contas

    lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
    If lErro <> SUCESSO Then Error 55772

    If iTestaLinhaAtual = NAO_TESTA_LINHA_ATUAL Then iLinhaAtual = GridContas.Row

    'Verifica se a conta ja esta presente no grid de contas
    For iLinha = 1 To objGridContas.iLinhasExistentes
    
        If iLinha <> iLinhaAtual Then
        
            lErro = CF("Conta_Formata", GridContas.TextMatrix(iLinha, iGrid_ContaInicio_Col), sContaInicioFormatadaGrid, iContaPreenchida)
            If lErro <> SUCESSO Then Error 55744
            
            lErro = CF("Conta_Formata", GridContas.TextMatrix(iLinha, iGrid_ContaFinal_Col), sContaFinalFormatadaGrid, iContaPreenchida1)
            If lErro <> SUCESSO Then Error 55745
                
            If iContaPreenchida = CONTA_PREENCHIDA And iContaPreenchida1 = CONTA_PREENCHIDA Then

                If sContaFormatada >= sContaInicioFormatadaGrid And sContaFormatada <= sContaFinalFormatadaGrid Then Error 55746
                
            End If
            
        End If
        
    Next

    Testa_Conta_Grid_Contas = SUCESSO
    
    Exit Function

Erro_Testa_Conta_Grid_Contas:

    Testa_Conta_Grid_Contas = Err

    Select Case Err

        Case 55744, 55745, 55772

        Case 55746
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_JA_UTILIZADA_GRID_CONTAS", Err, sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166131)

    End Select
    
    Exit Function

End Function

Private Function Testa_Conta_Grid_Lancamentos(sConta As String, iTestaLinhaAtual As Integer) As Long
'verifica se a conta passada como parametro coincide com as contas em lançamentos

Dim lErro As Long
Dim iLinha As Integer
Dim sContaFormatada As String
Dim sContaFormatadaGrid As String
Dim iContaPreenchida As Integer
Dim iLinhaAtual As Integer

On Error GoTo Erro_Testa_Conta_Grid_Lancamentos

    lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
    If lErro <> SUCESSO Then Error 55773

    If iTestaLinhaAtual = NAO_TESTA_LINHA_ATUAL Then iLinhaAtual = GridLancamentos.Row

    'Verifica se a conta ja esta presente no grid
    For iLinha = 1 To objGrid1.iLinhasExistentes
    
        If iLinha <> iLinhaAtual Then
    
            lErro = CF("Conta_Formata", GridLancamentos.TextMatrix(iLinha, iGrid_Conta_Col), sContaFormatadaGrid, iContaPreenchida)
            If lErro <> SUCESSO Then Error 55675
    
            If sContaFormatada = sContaFormatadaGrid Then Error 11278
            
        End If
        
    Next

    Testa_Conta_Grid_Lancamentos = SUCESSO
    
    Exit Function

Erro_Testa_Conta_Grid_Lancamentos:

    Testa_Conta_Grid_Lancamentos = Err

    Select Case Err

        Case 11278
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_JA_UTILIZADA_GRID_RATEIOS", Err, sConta)

        Case 55675, 55773

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166132)

    End Select
    
    Exit Function

End Function

Private Function Testa_Assoc_ContaCcl(sContaFormatada As String, objContaCcl As ClassContaCcl) As Long
'verifica se a conta passada como parametro tem associacao com o centro de custo em questao

Dim lErro As Long
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim sCcl As String

On Error GoTo Erro_Testa_Assoc_ContaCcl

    'se utiliza centro de custo extra-contabil
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
    
        'se o centro de custo foi preenchido
        If Len(GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Ccl_Col)) > 0 Then
        
            'verifica se a associação da conta com o centro de custo está cadastrado
            sCcl = GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Ccl_Col)
    
            lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then Error 20555
    
            objContaCcl.sConta = sContaFormatada
            objContaCcl.sCcl = sCclFormatada
    
            lErro = CF("ContaCcl_Le", objContaCcl)
            If lErro <> SUCESSO And lErro <> 5871 Then Error 20556
    
            'associação Conta x Centro de Custo/Lucro não cadastrada
            If lErro = 5871 Then Error 20557
            
        End If
        
    End If

    Testa_Assoc_ContaCcl = SUCESSO
    
    Exit Function

Erro_Testa_Assoc_ContaCcl:

    Testa_Assoc_ContaCcl = Err

    Select Case Err

        Case 20555, 20556, 20557

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166133)

    End Select
    
    Exit Function

End Function

Private Function Testa_Conta_Credito(sConta As String) As Long
'verifica se a conta passada como parametro coincide com a conta credito

Dim lErro As Long
Dim sContaCreditoFormatada As String
Dim iContaPreenchida As Integer
Dim sContaFormatada As String

On Error GoTo Erro_Testa_Conta_Credito

    lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
    If lErro <> SUCESSO Then Error 55774

    If Len(ContaCredito.ClipText) > 0 Then
    
        lErro = CF("Conta_Formata", ContaCredito.Text, sContaCreditoFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then Error 55673
        
        'Verifica se a conta foi utilizada como conta crédito
        If sContaFormatada = sContaCreditoFormatada Then Error 11274
    
    End If
          
    Testa_Conta_Credito = SUCESSO
    
    Exit Function

Erro_Testa_Conta_Credito:

    Testa_Conta_Credito = Err

    Select Case Err

        Case 11274
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_JA_UTILIZADA_CONTA_CREDITO", Err, sConta)

        Case 55675, 55774

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166134)

    End Select
    
    Exit Function

End Function

Private Sub BotaoGravar_Click()
    
Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 41534
    
    Call Limpa_Tela_RateioOff

    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 41534
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166135)
            
    End Select
    
    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long
            
Dim lErro As Long
Dim colRateioOff As New Collection
Dim colContas As New Collection
Dim objRateioOff As New ClassRateioOff
Dim iContaPreenchida As Integer
Dim iCclPreenchida As Integer
Dim sContaOrigem As String
Dim sContaCredito As String
Dim sCclOrigem As String
Dim sConta As String
Dim dTotal As Double

On Error GoTo Erro_Gravar_Registro
            
    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(Codigo.ClipText)) = 0 Then Error 14731
    
    objRateioOff.lCodigo = CLng(Codigo.Text)
    
    'Armazena o Tipo do Rateio
    objRateioOff.iTipo = TipoRateio.ItemData(TipoRateio.ListIndex)
    
    ' verifica se a cclorigem esta preenchida
    If Len(CclOrigem.ClipText) = 0 Then Error 11282
                
    lErro = CF("Ccl_Formata", CclOrigem.Text, sCclOrigem, iCclPreenchida)
    If lErro <> SUCESSO Then Error 18361
        
    objRateioOff.sCclOrigem = sCclOrigem
    
    objRateioOff.sDescricao = Descricao.Text
        
    'Verificar se a conta credito esta preechida
    If Len(ContaCredito.ClipText) = 0 Then Error 11285
    
    sConta = ContaCredito.Text
    
    lErro = CF("Conta_Formata", sConta, sContaCredito, iContaPreenchida)
    If lErro <> SUCESSO Then Error 18195
     
    objRateioOff.sContaCre = sContaCredito
      
    'Verifica se pelo menos uma linha do Grid está preenchida
    If objGrid1.iLinhasExistentes = 0 Then Error 11286
    
    dTotal = GridColuna_Soma(iGrid_Percentual_Col)

    'verifica se o percentual de rateio totaliza 100%
    If Abs(dTotal - 1) > DELTA_VALORMONETARIO2 Then Error 11287
      
    'Preenche a colRateioOff com as informacoes contidas no Grid
    lErro = Grid_RateioOff(colRateioOff)
    If lErro <> SUCESSO Then Error 11288
           
    'Preenche a colRateioOff com as informacoes contidas no Grid
    lErro = Grid_Contas(colContas)
    If lErro <> SUCESSO Then Error 55781
        
    lErro = CF("RateioOff_Grava", objRateioOff, colRateioOff, colContas)
    If lErro <> SUCESSO Then Error 11290
         
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 11282
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCLORIGEM_NAO_DIGITADO", Err)
        
        Case 11288, 11290, 18195, 18361, 55781
        
        Case 11285
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACREDITO_NAO_DIGITADA", Err)
        
        Case 11286
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_RATEIOOFF_GRAVAR", Err)
                       
        Case 11287
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOMA_PERCENTUAL_NAO_VALIDA", Err)
                
        Case 14731
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_DIGITADO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166136)
            
    End Select
    
    Exit Function
    
End Function

Private Function Grid_RateioOff(colRateioOff As Collection) As Long
'Armazena os dados do grid em colRateioOff

Dim iIndice1 As Integer
Dim objRateioOff As ClassRateioOff
Dim sConta As String
Dim sContaFormatada As String
Dim sCcl As String
Dim sCclFormatada As String
Dim iContaPreenchida As Integer
Dim iCclPreenchida As Integer
Dim lErro As Long

On Error GoTo Erro_Grid_RateioOff

    For iIndice1 = 1 To objGrid1.iLinhasExistentes
        
        Set objRateioOff = New ClassRateioOff
            
        objRateioOff.lSeq = iIndice1
  
        sConta = GridLancamentos.TextMatrix(iIndice1, iGrid_Conta_Col)
        
        If Len(Trim(sConta)) = 0 Then Error 11289
        
        lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then Error 11291
            
        'Armazena a conta
        objRateioOff.sConta = sContaFormatada
    
        'Armazena o percentual
        objRateioOff.dPercentual = CDbl(Format(GridLancamentos.TextMatrix(iIndice1, iGrid_Percentual_Col), "General Number"))
        
        If objRateioOff.dPercentual = 0 Then Error 11347
                        
        'Se está usando Centro de Custo/Lucro, armazena-o
        If iGrid_Ccl_Col <> 999 Then
                
            sCcl = GridLancamentos.TextMatrix(iIndice1, iGrid_Ccl_Col)
            
            lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then Error 20565
            
            If iCclPreenchida = CCL_PREENCHIDA Then
                objRateioOff.sCcl = sCclFormatada
            Else
                objRateioOff.sCcl = ""
            End If
                
        Else
            objRateioOff.sCcl = ""
        End If
                        
        'Armazena o objeto objRateioOff na coleção colRateioOff
        colRateioOff.Add objRateioOff
        
    Next
    
    Grid_RateioOff = SUCESSO

    Exit Function

Erro_Grid_RateioOff:

    Grid_RateioOff = Err

    Select Case Err
    
        Case 11289
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", Err)
    
        Case 11291, 20565
                
        Case 11347
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_INVALIDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166137)
            
    End Select
    
    Exit Function

End Function

Private Function Grid_Contas(colContas As Collection) As Long
'Armazena os dados do grid de contas em colContas

Dim iIndice1 As Integer
Dim sConta As String
Dim sContaFormatada As String
Dim sCcl As String
Dim sCclFormatada As String
Dim iContaPreenchida As Integer
Dim iCclPreenchida As Integer
Dim lErro As Long
Dim objRateioOffContas As ClassRateioOffContas

On Error GoTo Erro_Grid_Contas

    For iIndice1 = 1 To objGridContas.iLinhasExistentes
        
        Set objRateioOffContas = New ClassRateioOffContas
            
        sConta = GridContas.TextMatrix(iIndice1, iGrid_ContaInicio_Col)
        
        If Len(Trim(sConta)) = 0 Then Error 55782
        
        lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then Error 55783
            
        'Armazena a conta
        objRateioOffContas.sContaInicio = sContaFormatada
    
        sConta = GridContas.TextMatrix(iIndice1, iGrid_ContaFinal_Col)
        
        If Len(Trim(sConta)) = 0 Then Error 55784
        
        lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then Error 55785
            
        'Armazena a conta
        objRateioOffContas.sContaFim = sContaFormatada
                        
        If objRateioOffContas.sContaFim < objRateioOffContas.sContaInicio Then Error 55786
                        
        'Armazena o objeto objRateioOff na coleção colRateioOff
        colContas.Add objRateioOffContas
        
    Next
    
    Grid_Contas = SUCESSO

    Exit Function

Erro_Grid_Contas:

    Grid_Contas = Err

    Select Case Err
    
        Case 55782
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTAINICIO_GRIDCONTAS_NAO_INFORMADA", Err, iIndice1)
    
        Case 55783, 55785
                
        Case 55784
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTAFIM_GRIDCONTAS_NAO_INFORMADA", Err, iIndice1)
                
        Case 55786
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTAFIM_MENOR_CONTAINICIO", Err, iIndice1, GridContas.TextMatrix(iIndice1, iGrid_ContaInicio_Col), GridContas.TextMatrix(iIndice1, iGrid_ContaFinal_Col))
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166138)
            
    End Select
    
    Exit Function

End Function


Private Sub BotaoExcluir_Click()
'Exclui os lançamentos relativos ao Rateio digitado na tela
    
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lCodigo As Long

On Error GoTo Erro_BotaoExcluir_Click
     
    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Codigo.ClipText) = 0 Then Error 11292
     
    'Envia Mensagem pedindo confirmação da Exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RATEIO")
    
    If vbMsgRes = vbYes Then
        
        lCodigo = CLng(Codigo)
        
        'Exclui todos os lancamentos daquele Rateio automático
        lErro = CF("RateioOff_Exclui", lCodigo)
        If lErro <> SUCESSO Then Error 11295
    
        Call Limpa_Tela_RateioOff
        
        iAlterado = 0
                
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
               
        Case 11292
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RATEIOOFF_CODIGO_NAO_PREENCHIDO", Err)
            
        Case 11295
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166139)
        
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim iCodigo As Integer
Dim lErro As Long
Dim objRateioOff As New ClassRateioOff
Dim lCodigo As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 11298

    Call Limpa_Tela_RateioOff

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 11298

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166140)

    End Select

    Exit Sub
    
End Sub

Private Sub Conta_GotFocus()

Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_Conta_GotFocus

'    TvwContas.Visible = True
'    LabelContas.Visible = True
'    TvwCcls.Visible = False
'    LabelCcls.Visible = False
    BotaoConta.Tag = "Conta"
    
    ContaDescricao.Caption = ""

    Call Grid_Campo_Recebe_Foco(objGrid1)
      
    sConta = GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col)
    
    If Len(sConta) > 0 Then

        lErro = Conta_Exibe_Descricao(sConta)
        If lErro <> SUCESSO Then Error 11299
        
    End If
    
    Exit Sub
    
Erro_Conta_GotFocus:

    Select Case Err
    
        Case 11299
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166141)
    
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

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Private Sub GridLancamentos_RowColChange()

    Call Grid_RowColChange(objGrid1)
       
End Sub

Private Sub GridLancamentos_Scroll()

    Call Grid_Scroll(objGrid1)
    
End Sub

Sub Limpa_Tela_RateioOff()

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
            
    Call Grid_Limpa(objGrid1)
    Call Grid_Limpa(objGridContas)
    
    TotalPercentual.Caption = ""
    CclOrigem.PromptInclude = False
    CclOrigem.Text = ""
    CclOrigem.PromptInclude = True
    ContaCredito.PromptInclude = False
    ContaCredito.Text = ""
    ContaCredito.PromptInclude = True
    ContaDescricao.Caption = ""
    CclDescricao.Caption = ""
    Descricao.Text = ""
    Codigo.Text = ""
    ContaCreditoDescricao.Caption = ""
    ContaOrigemDescricao.Caption = ""
    CclOrigemDescricao.Caption = ""
    
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
    If lErro <> SUCESSO Then Error 11302

    'Verifica se a conta foi digitada
    If Len(sConta) > 0 Then
    
        'Busca os dados da conta
        lErro = CF("Conta_SelecionaUma", sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO And lErro <> 6030 Then Error 11303

        'se não encontrou a conta ==> erro
        If lErro = 6030 Then Error 11304
    
        'Coloca a descricao da conta na Tela
        ContaDescricao.Caption = objPlanoConta.sDescConta
    
    End If
    
    Conta_Exibe_Descricao = SUCESSO
    
    Exit Function

Erro_Conta_Exibe_Descricao:

    Conta_Exibe_Descricao = Err
    
    Select Case Err
    
        Case 11302, 11303
            
        Case 11304
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, sConta)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166142)
            
    End Select
        
    Exit Function

End Function

Function ContaOrigem_Exibe_Descricao(sConta As String) As Long
'exibe a descrição da conta no campo ContaOrigemDescricao. A conta passada como parametro deve estar mascarada

Dim sContaFormatada As String
Dim lErro As Long
Dim iContaPreenchida As Integer
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_ContaOrigem_Exibe_Descricao

    'Retorna conta formatada como no BD
    lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
    If lErro <> SUCESSO Then Error 55753

    'Verifica se a conta foi digitada
    If Len(sConta) > 0 Then
    
        'Busca os dados da conta
        lErro = CF("Conta_SelecionaUma", sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO And lErro <> 6030 Then Error 55754

        'se não encontrou a conta ==> erro
        If lErro = 6030 Then Error 55755
    
        'Coloca a descricao da conta na Tela
        ContaOrigemDescricao.Caption = objPlanoConta.sDescConta
    
    End If
    
    ContaOrigem_Exibe_Descricao = SUCESSO
    
    Exit Function

Erro_ContaOrigem_Exibe_Descricao:

    ContaOrigem_Exibe_Descricao = Err
    
    Select Case Err
    
        Case 55753, 55754
            
        Case 55755
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, sConta)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166143)
            
    End Select
        
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
    If lErro <> SUCESSO Then Error 24010
    
    For Each objCcl In colCcl
        
        sCclMascarado = String(STRING_CCL, 0)

        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then Error 24011

        If objCcl.iTipoCcl = CCL_ANALITICA Then
            sCcl = "A" & objCcl.sCcl
        Else
            sCcl = "S" & objCcl.sCcl
        End If

        sCclPai = String(STRING_CCL, 0)
        
        'retorna o centro de custo/lucro "pai" do centro de custo/lucro em questão, se houver
        lErro = Mascara_RetornaCclPai(objCcl.sCcl, sCclPai)
        If lErro <> SUCESSO Then Error 24012
        
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

        Case 24010
        
        Case 24011
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objCcl.sCcl)

        Case 24012
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCclPai", Err, objCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166144)

    End Select
    
    Exit Function

End Function

Private Sub TvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
    
Dim sCcl As String
Dim sCclMascarado As String
Dim lErro As Long
Dim sCclEnxuta As String
Dim lPosicaoSeparador As Long
Dim sCaracterInicial As String
Dim sCclFormatada As String
Dim objCcl As New ClassCcl
   
On Error GoTo Erro_TvwCcls_NodeClick
          
    If GridLancamentos.Col = iGrid_Ccl_Col And TvwCcls.Tag = "Ccl" Then
      
        sCaracterInicial = Left(Node.Key, 1)
    
        If sCaracterInicial = "A" Then
    
            sCcl = Right(Node.Key, Len(Node.Key) - 1)
              
            sCclEnxuta = String(STRING_CCL, 0)
            
            'volta mascarado apenas os caracteres preenchidos
            lErro = Mascara_RetornaCclEnxuta(sCcl, sCclEnxuta)
            If lErro <> SUCESSO Then Error 20570
            
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
    
    
    If TvwCcls.Tag = "CclOrigem" Then
        
        sCaracterInicial = Left(Node.Key, 1)
    
        If sCaracterInicial = "A" Then
    
            sCcl = Right(Node.Key, Len(Node.Key) - 1)
        
            sCclMascarado = String(STRING_CCL, 0)
            
            lErro = Mascara_MascararCcl(sCcl, sCclMascarado)
            If lErro <> SUCESSO Then Error 24013
            
            CclOrigem.PromptInclude = False
            CclOrigem.Text = sCclMascarado
            CclOrigem.PromptInclude = True
        
            'Preenche a Descricao do centro de custo/lucro
            lPosicaoSeparador = InStr(Node.Text, SEPARADOR)
            CclOrigemDescricao.Caption = Mid(Node.Text, lPosicaoSeparador + 1)
        
        End If
    End If
    
    Exit Sub

Erro_TvwCcls_NodeClick:

    Select Case Err
    
        Case 20570
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", Err, sCcl)
            
        Case 24013
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, sCcl)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166145)
            
    End Select
        
    Exit Sub
    
End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim colRateioOff As New Collection
Dim objRateioOff As New ClassRateioOff
Dim iContaPreenchida As Integer
Dim iCclPreenchida As Integer
Dim sContaOrigem As String
Dim sContaCredito As String
Dim sCclOrigem As String
Dim sConta As String
Dim dTotal As Double

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "RateioOffView"
        
    'Armazena op Tipo do Rateio
    If Len(Trim(TipoRateio.Text)) > 0 Then
        
        objRateioOff.iTipo = CInt(TipoRateio.ItemData(TipoRateio.ListIndex))
    
    Else
        
        objRateioOff.iTipo = 0
    
    End If
    
    'Armazena a Descrição
    If Len(Trim(Descricao.Text)) = 0 Then
       
       objRateioOff.sDescricao = String(STRING_RATEIO_DESCRICAO, 0)
    
    Else
       
       objRateioOff.sDescricao = Descricao.Text
     
    End If
     
     'Armazena o Código
    If Len(Trim(Codigo.Text)) = 0 Then
       
       objRateioOff.lCodigo = 0
    
    Else
       
       objRateioOff.lCodigo = Codigo.Text
     
     End If
     
    lErro = CF("Ccl_Formata", CclOrigem.Text, sCclOrigem, iCclPreenchida)
    If lErro <> SUCESSO Then Error 24152
        
    objRateioOff.sCclOrigem = sCclOrigem
        
    objRateioOff.sDescricao = Descricao.Text
        
    sConta = ContaCredito.Text
    
    lErro = CF("Conta_Formata", sConta, sContaCredito, iContaPreenchida)
    If lErro <> SUCESSO Then Error 24154
    
    objRateioOff.sContaCre = sContaCredito
      
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "CclOrigem", objRateioOff.sCclOrigem, STRING_CCL, "CclOrigem"
    colCampoValor.Add "Codigo", objRateioOff.lCodigo, 0, "Codigo"
    colCampoValor.Add "ContaCre", objRateioOff.sContaCre, STRING_CONTA, "ContaCre"
    colCampoValor.Add "Descricao", objRateioOff.sDescricao, STRING_RATEIO_DESCRICAO, "Descricao"
    colCampoValor.Add "Tipo", objRateioOff.iTipo, 0, "Tipo"
    
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case Err

        Case 24152, 24154
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166146)

    End Select
    
    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objRateioOff As New ClassRateioOff

On Error GoTo Erro_Tela_Preenche

    objRateioOff.lCodigo = colCampoValor.Item("Codigo").vValor

    If objRateioOff.lCodigo <> 0 Then
        
        objRateioOff.sDescricao = colCampoValor.Item("Descricao").vValor
        
        objRateioOff.sCclOrigem = colCampoValor.Item("CclOrigem").vValor
        
        lErro = Traz_Doc_Tela(objRateioOff)
        If lErro <> SUCESSO Then Error 24156

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 24156

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166147)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

'********************************
'Funções relativas ao GridContas
'********************************

Private Sub GridContas_Click()
Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridContas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridContas, iAlterado)
    End If

End Sub

Private Sub GridContas_EnterCell()
    Call Grid_Entrada_Celula(objGridContas, iAlterado)
End Sub

Private Sub GridContas_GotFocus()
    Call Grid_Recebe_Foco(objGridContas)
End Sub

Private Sub GridContas_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridContas)
End Sub

Private Sub GridContas_KeyPress(KeyAscii As Integer)
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridContas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridContas, iAlterado)
    End If

End Sub

Private Sub GridContas_LeaveCell()
    Call Saida_Celula(objGridContas)
End Sub

Private Sub GridContas_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridContas)
End Sub

Private Sub GridContas_RowColChange()
    Call Grid_RowColChange(objGridContas)
End Sub

Private Sub GridContas_Scroll()
    Call Grid_Scroll(objGridContas)
End Sub

Private Function Inicializa_Grid_Contas(objGridInt As AdmGrid) As Long
'Inicializa o grid de contas

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Contas

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Conta Início")
    objGridInt.colColuna.Add ("Conta Fim")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ContaInicio.Name)
    objGridInt.colCampo.Add (ContaFim.Name)

    'Colunas do Grid
    iGrid_ContaInicio_Col = 1
    iGrid_ContaFinal_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridContas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_CONTAS_RATEIOOFF + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridContas.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Contas = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Contas:

    Inicializa_Grid_Contas = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166148)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Contas(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Contas

    If objGridInt.objGrid Is GridContas Then

        Select Case GridContas.Col

            Case iGrid_ContaInicio_Col

                lErro = Saida_Celula_ContaInicio(objGridInt)
                If lErro <> SUCESSO Then Error 55731

            Case iGrid_ContaFinal_Col

                lErro = Saida_Celula_ContaFim(objGridInt)
                If lErro <> SUCESSO Then Error 55732

        End Select

    End If

    Saida_Celula_Contas = SUCESSO

    Exit Function

Erro_Saida_Celula_Contas:

    Saida_Celula_Contas = Err

    Select Case Err

        Case 55731, 55732

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166149)

    End Select

    Exit Function

End Function

Private Sub ContaInicio_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ContaInicio_GotFocus()

Dim lErro As Long
Dim sConta As String

On Error GoTo Erro_ContaInicio_GotFocus

'    TvwContas.Visible = True
'    LabelContas.Visible = True
'    TvwCcls.Visible = False
'    LabelCcls.Visible = False
    BotaoConta.Tag = "ContaInicio"
    
    
    ContaOrigemDescricao.Caption = ""

    Call Grid_Campo_Recebe_Foco(objGridContas)
      
    sConta = GridContas.TextMatrix(GridContas.Row, GridContas.Col)
    
    If Len(sConta) > 0 Then

        lErro = ContaOrigem_Exibe_Descricao(sConta)
        If lErro <> SUCESSO Then Error 55756
        
    End If
    
    Exit Sub
    
Erro_ContaInicio_GotFocus:

    Select Case Err
    
        Case 55756
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166150)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub ContaInicio_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContas)
End Sub

Private Sub ContaInicio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContas.objControle = ContaInicio
    lErro = Grid_Campo_Libera_Foco(objGridContas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ContaFim_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ContaFim_GotFocus()

Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_ContaFim_GotFocus

'    TvwContas.Visible = True
'    LabelContas.Visible = True
'    TvwCcls.Visible = False
'    LabelCcls.Visible = False
    BotaoConta.Tag = "ContaFim"
    
    
    ContaOrigemDescricao.Caption = ""

    Call Grid_Campo_Recebe_Foco(objGridContas)
      
    sConta = GridContas.TextMatrix(GridContas.Row, GridContas.Col)
    
    If Len(sConta) > 0 Then

        lErro = ContaOrigem_Exibe_Descricao(sConta)
        If lErro <> SUCESSO Then Error 55757
        
    End If
    
    Exit Sub
    
Erro_ContaFim_GotFocus:

    Select Case Err
    
        Case 55757
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166151)
    
    End Select
    
    Exit Sub

End Sub

Private Sub ContaFim_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContas)
End Sub

Private Sub ContaFim_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContas.objControle = ContaFim
    lErro = Grid_Campo_Libera_Foco(objGridContas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Function Saida_Celula_ContaInicio(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objPlanoConta As New ClassPlanoConta
Dim sContaMascarada As String

On Error GoTo Erro_Saida_Celula_ContaInicio

    Set objGridInt.objControle = ContaInicio

    If Len(Trim(ContaInicio.ClipText)) > 0 Then

        sContaFormatada = String(STRING_CONTA, 0)
        
        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica", ContaInicio.Text, ContaInicio.ClipText, objPlanoConta)
        If lErro <> SUCESSO And lErro <> 44033 And lErro <> 44037 Then Error 55777
        
        'se é uma conta simples, coloca a conta normal no lugar da conta simples
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then Error 55778
            
            ContaInicio.PromptInclude = False
            ContaInicio.Text = sContaMascarada
            ContaInicio.PromptInclude = True
            
        'se não encontrou a conta simples
        ElseIf lErro = 44033 Or lErro = 44037 Then
        
            'testa a conta no seu formato normal
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaInicio.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
            If lErro <> SUCESSO And lErro <> 5700 Then Error 55733
                        
            'conta não cadastrada
            If lErro = 5700 Then Error 55734
        
        End If
        
'        'verifica se a conta passada como parametro coincide com a conta credito
'        lErro = Testa_Conta_Credito(ContaInicio.Text)
'        If lErro <> SUCESSO Then Error 55764
'
'        'verifica se a conta passada como parametro coincide com as contas em lançamentos
'        lErro = Testa_Conta_Grid_Lancamentos(ContaInicio.Text, TESTA_LINHA_ATUAL)
'        If lErro <> SUCESSO Then Error 55765
'
'        'verifica se a conta passada como parametro coincide com as contas em lançamentos
'        lErro = Testa_Conta_Grid_Contas(ContaInicio.Text, NAO_TESTA_LINHA_ATUAL)
'        If lErro <> SUCESSO Then Error 55766
        
        If GridContas.Row - GridContas.FixedRows = objGridContas.iLinhasExistentes Then
            objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
        End If

        'Coloca a descricao daconta na tela
        ContaOrigemDescricao.Caption = objPlanoConta.sDescConta

    End If



    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 55735

    Saida_Celula_ContaInicio = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaInicio:

    Saida_Celula_ContaInicio = Err

    Select Case Err

        Case 55733, 55735, 55764, 55765, 55766, 55777
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 55734
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, ContaInicio.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 55778
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166152)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ContaFim(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objPlanoConta As New ClassPlanoConta
Dim sContaMascarada As String

On Error GoTo Erro_Saida_Celula_ContaFim

    Set objGridInt.objControle = ContaFim

    If Len(Trim(ContaFim.ClipText)) > 0 Then

        sContaFormatada = String(STRING_CONTA, 0)
        
        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica", ContaFim.Text, ContaFim.ClipText, objPlanoConta)
        If lErro <> SUCESSO And lErro <> 44033 And lErro <> 44037 Then Error 55779
        
        'se é uma conta simples, coloca a conta normal no lugar da conta simples
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then Error 55780
            
            ContaFim.PromptInclude = False
            ContaFim.Text = sContaMascarada
            ContaFim.PromptInclude = True
            
        'se não encontrou a conta simples
        ElseIf lErro = 44033 Or lErro = 44037 Then
        
            'testa a conta no seu formato normal
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaFim.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
            If lErro <> SUCESSO And lErro <> 5700 Then Error 55736
                        
            'conta não cadastrada
            If lErro = 5700 Then Error 55737
        
        End If

'        'verifica se a conta passada como parametro coincide com a conta credito
'        lErro = Testa_Conta_Credito(ContaFim.Text)
'        If lErro <> SUCESSO Then Error 55767
'
'        'verifica se a conta passada como parametro coincide com as contas em lançamentos
'        lErro = Testa_Conta_Grid_Lancamentos(ContaFim.Text, TESTA_LINHA_ATUAL)
'        If lErro <> SUCESSO Then Error 55768
'
'        'verifica se a conta passada como parametro coincide com as contas em lançamentos
'        lErro = Testa_Conta_Grid_Contas(ContaFim.Text, NAO_TESTA_LINHA_ATUAL)
'        If lErro <> SUCESSO Then Error 55769
        
        If GridContas.Row - GridContas.FixedRows = objGridContas.iLinhasExistentes Then
            objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
        End If

        'Coloca a descricao daconta na tela
        ContaOrigemDescricao.Caption = objPlanoConta.sDescConta

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 55738

    Saida_Celula_ContaFim = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaFim:

    Saida_Celula_ContaFim = Err

    Select Case Err

        Case 55736, 55738, 55767, 55768, 55769, 55777
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 55737
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, ContaFim.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 55780
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166153)

    End Select

    Exit Function

End Function
'********************************
' fim do tratamento do GridContas
'********************************


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    
    Parent.HelpContextID = IDH_RATEIO_OFF_LINE
    Set Form_Load_Ocx = Me
    Caption = "Rateio Off-Line"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RateioOff"
    
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
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcl_Click
        ElseIf Me.ActiveControl Is CclOrigem Then
            Call BotaoCcl_Click
        ElseIf Me.ActiveControl Is Conta Then
            Call BotaoConta_Click
        ElseIf Me.ActiveControl Is ContaCredito Then
            Call BotaoConta_Click
        ElseIf Me.ActiveControl Is ContaInicio Then
            Call BotaoConta_Click
        ElseIf Me.ActiveControl Is ContaFim Then
            Call BotaoConta_Click
        End If
    
    End If

End Sub


Private Sub ContaOrigemDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaOrigemDescricao, Source, X, Y)
End Sub

Private Sub ContaOrigemDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaOrigemDescricao, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub LabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCcl, Source, X, Y)
End Sub

Private Sub LabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCcl, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub CclOrigemDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclOrigemDescricao, Source, X, Y)
End Sub

Private Sub CclOrigemDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclOrigemDescricao, Button, Shift, X, Y)
End Sub

Private Sub CclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclLabel, Source, X, Y)
End Sub

Private Sub CclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclLabel, Button, Shift, X, Y)
End Sub

Private Sub CclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclDescricao, Source, X, Y)
End Sub

Private Sub CclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclDescricao, Button, Shift, X, Y)
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

Private Sub ContaCreditoDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaCreditoDescricao, Source, X, Y)
End Sub

Private Sub ContaCreditoDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaCreditoDescricao, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub TotalPercentual_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalPercentual, Source, X, Y)
End Sub

Private Sub TotalPercentual_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalPercentual, Button, Shift, X, Y)
End Sub

Private Sub LabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotais, Source, X, Y)
End Sub

Private Sub LabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotais, Button, Shift, X, Y)
End Sub

Private Sub LabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContas, Source, X, Y)
End Sub

Private Sub LabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContas, Button, Shift, X, Y)
End Sub

Private Sub LabelCcls_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCcls, Source, X, Y)
End Sub

Private Sub LabelCcls_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCcls, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub


Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Public Sub BotaoConta_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
    
    Select Case BotaoConta.Tag
    
        Case "ContaCredito"
    
            If Len(ContaCredito.Text) > 0 Then objPlanoConta.sConta = ContaCredito.Text
    
        Case "Conta"
    
            If Len(Conta.Text) > 0 Then objPlanoConta.sConta = Conta.Text
    
        Case "ContaInicio"
    
            If Len(ContaInicio.Text) > 0 Then objPlanoConta.sConta = ContaInicio.Text
    
    
        Case "ContaFim"
    
            If Len(ContaFim.Text) > 0 Then objPlanoConta.sConta = ContaFim.Text
    
    End Select
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("PlanoContaLista", colSelecao, objPlanoConta, objEventoConta)

End Sub

Private Sub objEventoConta_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sConta As String
Dim sContaEnxuta As String
Dim objContaCcl As New ClassContaCcl
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_objEventoConta_evSelecao
    
    Set objPlanoConta = obj1
    
    sConta = objPlanoConta.sConta
    
    'le a conta
    lErro = CF("PlanoConta_Le_Conta1", sConta, objPlanoConta)
    If lErro <> SUCESSO And lErro <> 6030 Then gError 197936
    
    If objPlanoConta.iAtivo <> CONTA_ATIVA Then gError 197937
    
    If objPlanoConta.iTipoConta <> CONTA_ANALITICA Then gError 197938
    
    sContaEnxuta = String(STRING_CONTA, 0)

    lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
    If lErro <> SUCESSO Then gError 197939

    Select Case BotaoConta.Tag
    
        Case "ContaCredito"
    
            ContaCredito.PromptInclude = False
            ContaCredito.Text = sContaEnxuta
            ContaCredito.PromptInclude = True
    
            ContaCreditoDescricao.Caption = objPlanoConta.sDescConta
    
        Case "Conta"
    
            If GridLancamentos.Col = iGrid_Conta_Col Then
            
                If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.iLinhasExistentes Then
                    objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
                End If
            
                'verifica se a conta passada como parametro tem associacao com o centro de custo em questao
'                lErro = Testa_Assoc_ContaCcl(sContaEnxuta, objContaCcl)
'                If lErro <> SUCESSO And lErro <> 20557 Then gError 197940
                
                'se está faltando a associacao da conta com o centro de custo
                If lErro = 20557 Then gError 197941

                Conta.PromptInclude = False
                Conta.Text = sContaEnxuta
                Conta.PromptInclude = True

                GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Conta_Col) = Conta.Text

                ContaDescricao.Caption = objPlanoConta.sDescConta

            End If
    
        Case "ContaInicio"
    
            If GridContas.Col = iGrid_ContaInicio_Col Then
            
                If objGridContas.objGrid.Row - objGridContas.objGrid.FixedRows = objGridContas.iLinhasExistentes Then
                    objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
                End If
            
                ContaInicio.PromptInclude = False
                ContaInicio.Text = sContaEnxuta
                ContaInicio.PromptInclude = True

                GridContas.TextMatrix(GridContas.Row, iGrid_ContaInicio_Col) = ContaInicio.Text

                ContaOrigemDescricao.Caption = objPlanoConta.sDescConta
                
            End If
    
        Case "ContaFim"
    
            If GridContas.Col = iGrid_ContaFinal_Col Then
            
                If objGridContas.objGrid.Row - objGridContas.objGrid.FixedRows = objGridContas.iLinhasExistentes Then
                    objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
                End If
            
                ContaFim.PromptInclude = False
                ContaFim.Text = sContaEnxuta
                ContaFim.PromptInclude = True

                GridContas.TextMatrix(GridContas.Row, iGrid_ContaFinal_Col) = ContaFim.Text

                ContaOrigemDescricao.Caption = objPlanoConta.sDescConta
                
            End If
    
    End Select

    Me.Show
    
    Exit Sub
    
Erro_objEventoConta_evSelecao:

    Select Case gErr
    
        Case 197936, 197940
    
        Case 197937
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_INATIVA", gErr, sConta)
        
        Case 197938
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_ANALITICA", gErr, sConta)
    
        Case 197939
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, sConta)
    
        Case 197941
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTACCL_INEXISTENTE", sConta, GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Ccl_Col))

            If vbMsgRes = vbYes Then
                Call Chama_Tela("ContaCcl", objContaCcl)
            End If
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197942)
        
    End Select

    Exit Sub

End Sub

Public Sub BotaoCcl_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection
Dim sCclOrigem As String
Dim iCclPreenchida As Integer
Dim lErro As Long

On Error GoTo Erro_LabelCcl_Click


    Select Case BotaoCcl.Tag
    
        Case "Ccl"
    
            If Len(Trim(Ccl.ClipText)) > 0 Then
            
                lErro = CF("Ccl_Formata", Ccl.Text, sCclOrigem, iCclPreenchida)
                If lErro <> SUCESSO Then gError 197943
        
                If iCclPreenchida = CCL_PREENCHIDA Then objCcl.sCcl = sCclOrigem
            Else
                objCcl.sCcl = ""
            End If
    
        Case "CclOrigem"
    
            'Verifica se o campo CclOrigem foi preenchido
            If Len(Trim(CclOrigem.ClipText)) > 0 Then
            
                lErro = CF("Ccl_Formata", CclOrigem.Text, sCclOrigem, iCclPreenchida)
                If lErro <> SUCESSO Then gError 197944
        
                If iCclPreenchida = CCL_PREENCHIDA Then objCcl.sCcl = sCclOrigem
            Else
                objCcl.sCcl = ""
            End If
    
    End Select

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)
    
    Exit Sub
    
Erro_LabelCcl_Click:

    Select Case gErr
        
        Case 197943, 197944
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197945)
            
    End Select

    Exit Sub

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objCcl As ClassCcl
Dim sConta As String
Dim sCclEnxuta As String

On Error GoTo Erro_objEventoCcl_evSelecao
    
    Set objCcl = obj1

    lErro = CF("Ccl_Le", objCcl)
    If lErro <> SUCESSO And lErro <> 5599 Then gError 197944

    If objCcl.iTipoCcl <> CCL_ANALITICA Then gError 197945
    
    If objCcl.iAtivo = 0 Then gError 197946
    
    sCclEnxuta = String(STRING_CONTA, 0)

    lErro = Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclEnxuta)
    If lErro <> SUCESSO Then gError 197947

    Select Case BotaoCcl.Tag

        Case "Ccl"

            If GridLancamentos.Col = iGrid_Ccl_Col Then

                Ccl.PromptInclude = False
                Ccl.Text = sCclEnxuta
                Ccl.PromptInclude = True
    
                GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col) = Ccl.Text
            
                If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.iLinhasExistentes Then
                    objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
                End If
    
                CclDescricao.Caption = objCcl.sDescCcl

            End If
            
        Case "CclOrigem"
        
            CclOrigem.PromptInclude = False
            CclOrigem.Text = sCclEnxuta
            CclOrigem.PromptInclude = True
        
            CclOrigemDescricao.Caption = objCcl.sDescCcl
        

    End Select

    Me.Show
    
    Exit Sub
    
Erro_objEventoCcl_evSelecao:

    Select Case gErr
    
        Case 197944

        Case 197945
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_ANALITICA1", gErr, objCcl.sCcl)
  
        Case 197946
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_INATIVO", gErr, objCcl.sCcl)

        Case 197947
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCcl)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197948)
        
    End Select

    Exit Sub

End Sub


