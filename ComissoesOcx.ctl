VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl ComissoesOcx 
   ClientHeight    =   7350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9525
   KeyPreview      =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   9525
   Begin VB.Frame Frame2 
      Caption         =   "Informações do Documento"
      Height          =   855
      Left            =   120
      TabIndex        =   35
      Top             =   2520
      Width           =   9255
      Begin VB.Label Valor 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7215
         TabIndex        =   41
         Top             =   360
         Width           =   1905
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
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
         Left            =   6615
         TabIndex        =   40
         Top             =   390
         Width           =   510
      End
      Begin VB.Label Filial 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4080
         TabIndex        =   39
         Top             =   360
         Width           =   1905
      End
      Begin VB.Label Cliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   975
         TabIndex        =   38
         Top             =   360
         Width           =   2145
      End
      Begin VB.Label Label3 
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
         Left            =   3525
         TabIndex        =   37
         Top             =   390
         Width           =   465
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
         Left            =   240
         TabIndex        =   36
         Top             =   390
         Width           =   660
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   795
      Left            =   6360
      ScaleHeight     =   735
      ScaleWidth      =   3030
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   3090
      Begin VB.CommandButton BotaoFechar 
         Height          =   600
         Left            =   2520
         Picture         =   "ComissoesOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   600
         Left            =   2025
         Picture         =   "ComissoesOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   600
         Left            =   1530
         Picture         =   "ComissoesOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoConsultar 
         Height          =   600
         Left            =   120
         Picture         =   "ComissoesOcx.ctx":080A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   75
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Documento a ser Consultado"
      Height          =   1335
      Left            =   120
      TabIndex        =   23
      Top             =   960
      Width           =   9255
      Begin VB.OptionButton OptionCupom 
         Caption         =   "Cupom Fiscal"
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
         Left            =   765
         TabIndex        =   2
         Top             =   1035
         Width           =   1800
      End
      Begin VB.OptionButton PedidoVenda 
         Caption         =   "Pedido de Venda"
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
         Left            =   750
         TabIndex        =   0
         Top             =   315
         Width           =   1800
      End
      Begin VB.ComboBox Serie 
         Height          =   315
         Left            =   7485
         TabIndex        =   5
         Top             =   255
         Visible         =   0   'False
         Width           =   765
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   4800
         TabIndex        =   3
         Top             =   270
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin VB.OptionButton NotaFiscal 
         Caption         =   "Nota Fiscal"
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
         Left            =   750
         TabIndex        =   1
         Top             =   675
         Width           =   1485
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   315
         Left            =   4800
         TabIndex        =   7
         Top             =   795
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox COO 
         Height          =   300
         Left            =   4800
         TabIndex        =   4
         Top             =   270
         Visible         =   0   'False
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ECF 
         Height          =   315
         Left            =   7485
         TabIndex        =   6
         Top             =   255
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label LabelECF 
         AutoSize        =   -1  'True
         Caption         =   "ECF:"
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
         Left            =   6975
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   45
         Top             =   315
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label LabelCOO 
         AutoSize        =   -1  'True
         Caption         =   "COO:"
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
         Left            =   4230
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   44
         Top             =   315
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Data Emissão:"
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
         Left            =   3480
         TabIndex        =   27
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label LabelSerie 
         AutoSize        =   -1  'True
         Caption         =   "Série:"
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
         Left            =   6885
         TabIndex        =   25
         Top             =   300
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label LabelNumero 
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
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3960
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   26
         Top             =   315
         Width           =   735
      End
   End
   Begin VB.Frame SSFrame4 
      Caption         =   "Comissões"
      Height          =   3660
      Index           =   0
      Left            =   90
      TabIndex        =   24
      Top             =   3570
      Width           =   9285
      Begin VB.CommandButton BotaoConsultaDocumento 
         Caption         =   "Consulta Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   6615
         Picture         =   "ComissoesOcx.ctx":25CC
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2535
         Width           =   1305
      End
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
         Height          =   870
         Left            =   7950
         Picture         =   "ComissoesOcx.ctx":309A
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2520
         Width           =   1185
      End
      Begin VB.Frame SSFrame4 
         Caption         =   "Totais - Comissões"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   155
         TabIndex        =   28
         Top             =   2460
         Width           =   6255
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor Total:"
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
            Index           =   36
            Left            =   120
            TabIndex        =   43
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label TotalValorBase 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1200
            TabIndex        =   42
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Percentual:"
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
            Index           =   19
            Left            =   2520
            TabIndex        =   32
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
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
            Index           =   20
            Left            =   4560
            TabIndex        =   31
            Top             =   360
            Width           =   615
         End
         Begin VB.Label TotalValorComissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   5160
            TabIndex        =   30
            Top             =   360
            Width           =   975
         End
         Begin VB.Label TotalPercentualComissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3600
            TabIndex        =   29
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.ComboBox DiretoIndireto 
         Height          =   315
         ItemData        =   "ComissoesOcx.ctx":3644
         Left            =   7680
         List            =   "ComissoesOcx.ctx":364E
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin MSMask.MaskEdBox PercentualComissao 
         Height          =   225
         Left            =   1740
         TabIndex        =   9
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorComissao 
         Height          =   225
         Left            =   3765
         TabIndex        =   11
         Top             =   615
         Width           =   1110
         _ExtentX        =   1958
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
      Begin MSMask.MaskEdBox ValorBase 
         Height          =   225
         Left            =   2700
         TabIndex        =   10
         Top             =   585
         Width           =   1170
         _ExtentX        =   2064
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
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   225
         Left            =   750
         TabIndex        =   8
         Top             =   540
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorEmissao 
         Height          =   225
         Left            =   5700
         TabIndex        =   13
         Top             =   600
         Width           =   1275
         _ExtentX        =   2249
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
      Begin MSMask.MaskEdBox PercentualEmissao 
         Height          =   225
         Left            =   4770
         TabIndex        =   12
         Top             =   615
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorBaixa 
         Height          =   225
         Left            =   7815
         TabIndex        =   15
         Top             =   585
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   397
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
      Begin MSMask.MaskEdBox PercentualBaixa 
         Height          =   225
         Left            =   6900
         TabIndex        =   14
         Top             =   585
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0%"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridComissoes 
         Height          =   1845
         Left            =   60
         TabIndex        =   17
         Top             =   330
         Width           =   9090
         _ExtentX        =   16034
         _ExtentY        =   3254
         _Version        =   393216
         Rows            =   7
         Cols            =   5
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
      End
   End
End
Attribute VB_Name = "ComissoesOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'VARIAVEIS GLOBAIS DA TELA
Dim iAlterado As Integer
Dim iTipo_Documento As Integer
Dim objGrid1 As AdmGrid

'EVENTOS DE BROWSER
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoCodigoNF As AdmEvento
Attribute objEventoCodigoNF.VB_VarHelpID = -1
Private WithEvents objEventoCodigoPV As AdmEvento
Attribute objEventoCodigoPV.VB_VarHelpID = -1
Private WithEvents objEventoCupomFiscal As AdmEvento
Attribute objEventoCupomFiscal.VB_VarHelpID = -1

'Numero de linhas do grid
'tulio 9/5/02
Const NUM_MAX_COMISSOES = 100

'Campos do Grid
Dim iGrid_Vendedor_Col As Integer
Dim iGrid_Percentual_Comissao_Col As Integer
Dim iGrid_Valor_Base_Col As Integer
Dim iGrid_Valor_Comissao_Col As Integer
Dim iGrid_Percentual_Emissao_Col As Integer
Dim iGrid_Valor_Emissao_Col As Integer
Dim iGrid_Percentual_Baixa_Col As Integer
Dim iGrid_Valor_Baixa_Col As Integer
Dim iGrid_DiretoIndireto_Col As Integer

'mario
'inicia objeto associado a GridComissoes
Public objTabComissoes As New ClassTabComissoes
Public objGridComissoes As AdmGrid

'******************************************
'4 eventos do controle do Grid de Comissoes: DiretoIndireto
'tulio 9/5/02
'******************************************

Private Sub DiretoIndireto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DiretoIndireto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub DiretoIndireto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub DiretoIndireto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = DiretoIndireto
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataEmissao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)

End Sub

Public Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a data de emissao foi digitada
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then gError 89997

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True

    Select Case gErr

        Case 89997

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154378)

    End Select

    Exit Sub

End Sub

Private Sub BotaoConsultar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoConsultar_Click

    'Caso seja Pedido de Venda, verifica se o Codigo está preenchido
    If Len(Trim(Codigo.Text)) = 0 And PedidoVenda.Value = True Then gError 43665
    
    'Caso seja uma Nota Fiscal, verifica se o Número, Série estão Preenchidos
    If (Len(Trim(Codigo.Text)) = 0 Or Len(Trim(Serie.Text)) = 0) And Serie.Visible = True Then gError 43666
    
    'Caso seja um Pedido
    If PedidoVenda.Value = True Then
        
        'Preenche o Grid com a comissao do Pedido
        lErro = PedidoDeVenda_PreencheGrid()
        If lErro <> SUCESSO Then gError 43663

    ElseIf NotaFiscal.Value = True Then
        
        'Caso seja uma Nota Fiscal
        If Len(Trim(Serie.Text)) <> 0 Then
            
            'Preenche o Grid com as Comissao da NFiscal
            lErro = NotaFisc_PreencheGrid()
            If lErro <> SUCESSO Then gError 43664

        End If

    Else
    
        'Preenche o Grid com a comissao associada ao cupom fiscal
        lErro = CupomFiscal_PreencheGrid()
        If lErro <> SUCESSO Then gError 126312

    End If
    
    Exit Sub
    
Erro_BotaoConsultar_Click:

    Select Case gErr
    
        Case 43663, 43664, 126312

        Case 43665
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_INFORMADO", gErr)

        Case 43666
            Call Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_NAO_INFORMADA", gErr)

        Case 126311

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154379)

    End Select

    Exit Sub

End Sub

Private Sub BotaoConsultaDocumento_Click()

Dim lErro As Long
Dim objPedidoDeVenda As New ClassPedidoDeVenda
Dim objNFiscal As New ClassNFiscal
Dim sTela As String

On Error GoTo Erro_BotaoConsultaDocumento_Click
    
    'Caso seja uma NFiscal
    If NotaFiscal.Value = True Then
        
        'Verifica se o Número e a Série estão preenchidos
        If Len(Trim(Codigo.Text)) = 0 Or Len(Trim(Serie.Text)) = 0 Then gError 43661
        
        'Preenche objNFiscal
        objNFiscal.sSerie = Serie.Text
        objNFiscal.lNumNotaFiscal = CLng(Codigo.Text)
        objNFiscal.iFilialEmpresa = giFilialEmpresa
        
        'Lê o Nome da Tela com a Série e o Número passados
        lErro = CF("TipoDocInfo_Le_NomeTela_NFiscal", objNFiscal, sTela)
        If lErro <> SUCESSO And lErro <> 58180 Then gError 58175
        
        'Se não encontrar ---> Erro
        If lErro = 58180 Then gError 58181
        
        'Chama a Tela Correspondente a NFiscal
        Call Chama_Tela(sTela, objNFiscal)
    
    End If
    
    'Caso seja um Pedido de Venda
    If PedidoVenda.Value = True Then
    
        If Len(Trim(Codigo.Text)) = 0 Then gError 43662
        
        'Passa para objPedidoDeVenda
        objPedidoDeVenda.lCodigo = CLng(Codigo.Text)
        objPedidoDeVenda.iFilialEmpresa = giFilialEmpresa
        
        'Chama a tela de Pedido de Venda
        Call Chama_Tela("PedidoVenda", objPedidoDeVenda)
    
    End If

    If OptionCupom.Value = True Then gError 126328

    Exit Sub

Erro_BotaoConsultaDocumento_Click:

    Select Case gErr

        Case 43661
            Call Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_NAO_INFORMADA", gErr)

        Case 43662
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_INFORMADO", gErr)
        
        Case 58175
        
        Case 58181
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", gErr, objNFiscal.lNumNotaFiscal)
        
        Case 126328
            Call Rotina_Erro(vbOKOnly, "ERRO_CUPOM_FISCAL_SEM_TELA_EDICAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154380)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    'Grava a Comissão
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 21341

    Call Limpa_Tela_Comissoes

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 21341

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154381)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objSerie As New ClassSerie
Dim lTamanho As Long
Dim sVendedor As String

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Codigo foi preenchido
    If iTipo_Documento <> CUPOM_FISCAL Then
        If Len(Codigo.Text) = 0 Then gError 21342
    End If
    
    If objGrid1.iLinhasExistentes > 0 Then

        'Loop de Validação dos dados do GridComissoesEmissao
        For iIndice = 1 To objGrid1.iLinhasExistentes
            
            'Verifica se o Vendedor foi informado
            If Len(Trim(GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col))) = 0 Then gError 21453
                
            'Verifica se o Percentual foi informado
            lTamanho = Len(Trim(GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Comissao_Col)))
            If lTamanho = 0 Then gError 21454

            'Verifica se Valor Base foi digitado
            If Len(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Base_Col)) = 0 Then gError 21455

            'Verifica se o Percentual de emissão foi informado
            lTamanho = Len(Trim(GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Emissao_Col)))
            If lTamanho = 0 Then gError 21456

            'Verifica se Valor Emissao foi informado
            If Len(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Emissao_Col)) = 0 Then gError 21457

        Next
        
    End If
    
    'Caso seja Pedido de Venda
    If iTipo_Documento = PEDIDO_DE_VENDA Then
        
        'Chama a rotina de Gravacao de Comissão para Pedido de Venda
        lErro = Grava_Registro_PedidoDeVenda()
        If lErro <> SUCESSO Then gError 21343

    ElseIf iTipo_Documento = NOTA_FISCAL Then
    'Caso seja uma NFiscal
        
        'Verifica se a série está Preenchida
        If Len(Trim(Serie.Text)) = 0 Then gError 21447
        
        'Preenche o objSerie
        objSerie.sSerie = Serie.Text
        
        'Le a serie no BD
        lErro = CF("Serie_Le", objSerie)
        If lErro <> SUCESSO And lErro <> 22202 Then gError 21448
        
        'Se não encontrar ---> EERO
        If lErro = 22202 Then gError 21344
        
        If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 89998
        
        'Chama Rotina de Gravacao de Comissão para NFiscal
        lErro = Grava_Registro_NFsRec()
        If lErro <> SUCESSO Then gError 21345

    ElseIf iTipo_Documento = CUPOM_FISCAL Then

        'Chama a rotina de Gravacao de Comissão para Cupom Fiscal
        lErro = Grava_Registro_CupomFiscal()
        If lErro <> SUCESSO Then gError 21343


    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 21342
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 21344
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, Serie.Text)

        Case 21447
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 21343, 21345, 21448 'Tratados nas rotinas chamadas

        Case 21453
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_COMISSAO_GRID_NAO_INFORMADO", gErr, iIndice)
        
        Case 21454
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_COMISSAO_NAO_INFORMADO", gErr, iIndice)

        Case 21455
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORBASE_COMISSAO_NAO_INFORMADO", gErr, iIndice)

        Case 21456
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_EMISSAO_NAO_INFORMADO", gErr, iIndice)

        Case 21457
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOREMISSAO_COMISSAO_NAO_INFORMADO", gErr, iIndice)
        
        Case 89998
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154382)

    End Select

    Exit Function

End Function

Function Grava_Registro_PedidoDeVenda() As Long
'Grava a Comissão do Pedido de Venda

Dim objPedVendas As New ClassPedidoDeVenda
Dim objComissaoPV As ClassComissaoPedVendas
Dim objVendedor As New ClassVendedor
Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim iCodigo As Integer

On Error GoTo Erro_Grava_Registro_PedidoDeVenda

    'Preenche objPedVendas
    objPedVendas.lCodigo = StrParaLong(Codigo.Text)
    objPedVendas.iFilialEmpresa = giFilialEmpresa
    
    'Passa o Codigo do Vendedor, os Valores e Percentuais para a Coleção de Comissoes de objPedVendas
    'Para Cada item do Grid
    For iIndice = 1 To objGrid1.iLinhasExistentes
        
        'Instancia um novo objComissaoPV
        Set objComissaoPV = New ClassComissaoPedVendas
        
        objVendedor.sNomeReduzido = GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col)
 
        'Lê o código do Vendedor
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then Error 64132
        
        'Se não encontrar -- > erro
        If lErro = 25008 Then Error 21346
        
        'Guarda no obj os dados que serão adicionados à coleção
        With objComissaoPV
        
            .iCodVendedor = objVendedor.iCodigo
            .dValorBase = StrParaDbl(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Base_Col))
            .dPercentual = PercentParaDbl(GridComissoes.TextMatrix(iIndice, GRID_PERCENTUAL_COL))
            .dValor = StrParaDbl(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Comissao_Col))
            .dPercentualEmissao = PercentParaDbl(GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Emissao_Col))
            .dValorEmissao = StrParaDbl(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Emissao_Col))
        
            'tulio 9/5/02
             If GridComissoes.TextMatrix(iIndice, iGrid_DiretoIndireto_Col) = VENDEDOR_DIRETO_STRING Then
            
                .iIndireta = VENDEDOR_DIRETO
                
            Else
            
                .iIndireta = VENDEDOR_INDIRETO
                
            End If
            
            .iSeq = iIndice
        
        End With
        
        'Adiciona na Coleção de Comissões
        objPedVendas.colComissoes.Add objComissaoPV

    Next
    
    'Grava a Comissão de Pedido de Venda
    lErro = CF("PedidoDeVenda_Grava_Comissoes", objPedVendas)
    If lErro <> SUCESSO Then Error 21347

    Grava_Registro_PedidoDeVenda = SUCESSO

    Exit Function

Erro_Grava_Registro_PedidoDeVenda:

    Grava_Registro_PedidoDeVenda = Err

    Select Case Err

        Case 21346
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", Err, objVendedor.sNomeReduzido)

        Case 21347, 64132 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154383)

    End Select

    Exit Function

End Function

Function Grava_Registro_NFsRec() As Long
'Grava a Comissão de Nota Fiscal

Dim objNFiscal As New ClassNFiscal
Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim objComissaoNF As ClassComissaoNF
Dim objVendedor As New ClassVendedor
Dim objcliente As New ClassCliente

On Error GoTo Erro_Grava_Registro_NFsRec

    objNFiscal.sSerie = String(STRING_SERIE, 0)
    
    'preenche o objNFiscal
    objNFiscal.lNumNotaFiscal = StrParaLong(Codigo.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.sSerie = Serie.Text
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.dValorTotal = StrParaDbl(Valor.Caption)
   
    'Passa o Cod do Vededor, Valores e Percentuais para a Coleção
    'Para cada item do Grid de Comissões
    For iIndice = 1 To objGrid1.iLinhasExistentes
        
        'Instancia um novo objComissaoNF
        Set objComissaoNF = New ClassComissaoNF
        
        objVendedor.sNomeReduzido = GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col)
 
        'Lê o código do Vendedor
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then gError 64133
        
        'Se não encontrou  ---> ERRO
        If lErro <> SUCESSO Then gError 21348
        
        'Guarda os dados no obj
        With objComissaoNF
        
            .iCodVendedor = objVendedor.iCodigo
            .dValorBase = StrParaDbl(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Base_Col))
            .dPercentual = PercentParaDbl(GridComissoes.TextMatrix(iIndice, GRID_PERCENTUAL_COL))
            .dValor = StrParaDbl(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Comissao_Col))
            .dPercentualEmissao = PercentParaDbl(GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Emissao_Col))
            .dValorEmissao = StrParaDbl(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Emissao_Col))
            
            'tulio 9/5/02
            If GridComissoes.TextMatrix(iIndice, iGrid_DiretoIndireto_Col) = VENDEDOR_DIRETO_STRING Then
            
                .iIndireta = VENDEDOR_DIRETO
                
            Else
            
                .iIndireta = VENDEDOR_INDIRETO
                
            End If

            .iSeq = iIndice
            
            'Passa para a coleção de Comissoes de objNfiscal
            objNFiscal.ColComissoesNF.Add objComissaoNF
        
        End With

    Next
    
    If Len(Trim(Cliente.Caption)) = 0 Then gError 209072
    
    objcliente.sNomeReduzido = Cliente.Caption
    
    lErro = CF("Cliente_Le_NomeReduzido", objcliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 89995
    
    'se o nome reduzido não estiver cadastrado ==> erro
    If lErro = 12348 Then gError 89996
    
    objNFiscal.iFilialCli = Codigo_Extrai(Filial.Caption)
    objNFiscal.lCliente = objcliente.lCodigo
    
    'Grava a Comissãqo da Nota Fiscal
    lErro = CF("NFiscal_Grava_Comissoes", objNFiscal)
    If lErro <> SUCESSO Then Error 21349

    Grava_Registro_NFsRec = SUCESSO

    Exit Function

Erro_Grava_Registro_NFsRec:

    Grava_Registro_NFsRec = gErr

    Select Case gErr

        Case 21348
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", gErr, GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col))

        Case 21349, 64133, 89995

        Case 89996
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Caption)

        Case 209072
            Call Rotina_Erro(vbOKOnly, "ERRO_TRAZER_NF_TELA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154384)

    End Select

    Exit Function

End Function

Function Grava_Registro_CupomFiscal() As Long
'Grava a Comissão do Cupom Fiscal

Dim objCupomFiscal As New ClassCupomFiscal
Dim objComissoesCF As ClassComissoesCF
Dim objVendedor As New ClassVendedor
Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim iCodigo As Integer
Dim colComissao As New Collection

On Error GoTo Erro_Grava_Registro_CupomFiscal

    'Preenche objPedVendas
    objCupomFiscal.lNumero = StrParaLong(COO.Text)
    objCupomFiscal.iFilialEmpresa = giFilialEmpresa
    objCupomFiscal.iECF = StrParaInt(ECF.Text)
    
    lErro = CF("CupomFiscal_Le", objCupomFiscal)
    If lErro <> SUCESSO And lErro <> 105262 Then gError 126343
    
    'se o cupom nao estiver cadastrado ==> erro
    If lErro = 105262 Then gError 126344
    
    'Passa o Codigo do Vendedor, os Valores e Percentuais para a Coleção de Comissoes de objPedVendas
    'Para Cada item do Grid
    For iIndice = 1 To objGrid1.iLinhasExistentes
        
        'Instancia um novo objComissaoPV
        Set objComissoesCF = New ClassComissoesCF
        
        objVendedor.sNomeReduzido = GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col)
 
        'Lê o código do Vendedor
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then gError 126345
        
        'Se não encontrar -- > erro
        If lErro = 25008 Then Error 126346
        
        'Guarda no obj os dados que serão adicionados à coleção
        With objComissoesCF
        
            .iCodVendedor = objVendedor.iCodigo
            .dValorBase = CDbl(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Base_Col))
            .dValorComissao = CDbl(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Comissao_Col))
        
        
            'tulio 9/5/02
             If GridComissoes.TextMatrix(iIndice, iGrid_DiretoIndireto_Col) = VENDEDOR_DIRETO_STRING Then
            
                .iIndireta = VENDEDOR_DIRETO
                
            Else
            
                .iIndireta = VENDEDOR_INDIRETO
                
            End If

        
        
        End With
        
        'Adiciona na Coleção de Comissões
        colComissao.Add objComissoesCF

    Next
    
    'Grava a Comissão de Pedido de Venda
    lErro = CF("Comissoes_Gravar_Loja_1", objCupomFiscal, colComissao)
    If lErro <> SUCESSO Then gError 126357
    
    Grava_Registro_CupomFiscal = SUCESSO

    Exit Function

Erro_Grava_Registro_CupomFiscal:

    Grava_Registro_CupomFiscal = gErr

    Select Case gErr

        Case 126343, 126345, 126357

        Case 126344
            Call Rotina_Erro(vbOKOnly, "ERRO_CUPOM_FISCAL_NAO_CADASTRADO2", gErr, objCupomFiscal.lNumero, objCupomFiscal.iFilialEmpresa, objCupomFiscal.iECF)

        Case 126346
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", gErr, objVendedor.sNomeReduzido)


        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154385)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO And lErro <> 20323 Then Error 21350
    
    Call Limpa_Tela_Comissoes

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 21350

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154386)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_Comissoes()

    Cliente.Caption = ""
    Filial.Caption = ""
    Valor.Caption = ""
    TotalPercentualComissao.Caption = ""
    TotalValorComissao.Caption = ""
    TotalValorBase.Caption = ""
    Serie.Text = ""

    'Limpa MaskedEditBox
    Call Limpa_Tela(Me)

    'Limpa o Grid
    Call Grid_Limpa(objGrid1)

    iAlterado = 0

End Sub

Private Sub BotaoVendedores_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection
    
    'Chama tela que lista todos os vendores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoVendedor = Nothing
    Set objEventoCodigoNF = Nothing
    Set objEventoCodigoPV = Nothing
    Set objEventoCupomFiscal = Nothing

    Set objGrid1 = Nothing
    
'mario
    Set objTabComissoes = Nothing
    
    
End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objPedidoDeVenda As New ClassPedidoDeVenda

On Error GoTo Erro_Codigo_Validate
    
    'Se a Nota Fiscal estiver ativa e Série e Codigo estiverem Preenchidos
    If NotaFiscal.Value = True And Len(Trim(Serie.Text)) > 0 And Len(Trim(Codigo.Text)) > 0 Then
            
        objNFiscal.lNumNotaFiscal = CLng(Codigo.Text)
        objNFiscal.sSerie = LTrim(Serie.Text)
        objNFiscal.iFilialEmpresa = giFilialEmpresa
        
        'Verifica se está Preenchido no BD
        lErro = CF("NF_NFFatura_Le_NumeroSerie", objNFiscal)
        If lErro <> SUCESSO And lErro <> 58324 Then Error 58330
        
        If lErro = 58324 Then Error 58331
    
    End If
    
    'Se for o Pedido que está ativo
    If PedidoVenda.Value = True And Len(Trim(Codigo.Text)) > 0 Then
        
        objPedidoDeVenda.lCodigo = CLng(Codigo.Text)
        objPedidoDeVenda.iFilialEmpresa = giFilialEmpresa
        
        'Lê o Pedido de Venda
        lErro = CF("PedidoDeVenda_Le", objPedidoDeVenda)
        If lErro <> SUCESSO And lErro <> 26509 Then Error 58390

        'Não achou o Pedido de Venda --> ERRO
        If lErro = 26509 Then Error 58391
        
    End If
    
    Exit Sub
    
Erro_Codigo_Validate:

    Cancel = True
    
    Select Case Err
        
        Case 58330 'Tratado na Rotina chamada
        
        Case 58331
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_TEM_COMISSAO", Err, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.iFilialEmpresa)
            
        Case 58391
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO1", Err, objPedidoDeVenda.lCodigo, objPedidoDeVenda.iFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154387)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelCOO_Click()

Dim objCupomFiscal As New ClassCupomFiscal
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCOO_Click

    'se o ECF estiver preenchido
    If Len(Trim(ECF.Text)) > 0 Then

        'move o ECF para o obj
        objCupomFiscal.iECF = StrParaInt(ECF.Text)

    End If

    'se o COO estiver preenchido
    If Len(Trim(COO.Text)) > 0 Then

        'move o COO para o obj
        objCupomFiscal.lNumero = StrParaLong(COO.Text)

    End If

    'Chama o Browser '
    Call Chama_Tela("CupomFiscalLista", colSelecao, objCupomFiscal, objEventoCupomFiscal)

    Exit Sub

Erro_LabelCOO_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154388)

    End Select

    Exit Sub

End Sub

Private Sub LabelECF_Click()

Dim objCupomFiscal As New ClassCupomFiscal
Dim colSelecao As New Collection

On Error GoTo Erro_LabelECF_Click

    'se o ECF estiver preenchido
    If Len(Trim(ECF.Text)) > 0 Then

        'move o ECF para o obj
        objCupomFiscal.iECF = StrParaInt(ECF.Text)

    End If

    'se o COO estiver preenchido
    If Len(Trim(COO.Text)) > 0 Then

        'move o COO para o obj
        objCupomFiscal.lNumero = StrParaLong(COO.Text)

    End If

    'Chama o Browser '
    Call Chama_Tela("CupomFiscalLista", colSelecao, objCupomFiscal, objEventoCupomFiscal)

    Exit Sub

Erro_LabelECF_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154389)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCupomFiscal_evSelecao(obj1 As Object)

Dim objCupomFiscal As ClassCupomFiscal
Dim lErro As Long

On Error GoTo Erro_objEventoCupomFiscal_evSelecao

    Set objCupomFiscal = obj1

    Call Limpa_Tela_Comissoes

    ECF.Text = CStr(objCupomFiscal.iECF)
    
    COO.Text = CStr(objCupomFiscal.lNumero)

    'Preenche o Grid de Comissões
    lErro = CupomFiscal_PreencheGrid
    If lErro <> SUCESSO Then gError 126349

    Me.Show

    Exit Sub

Erro_objEventoCupomFiscal_evSelecao:

    Select Case gErr
        
        Case 126349
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154390)

    End Select

    Exit Sub

End Sub

Private Sub LabelNumero_Click()

Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objNFiscal As New ClassNFiscal
Dim colSelecao As New Collection

    If PedidoVenda.Value = True Then
        
        'Chama a tela que lista todos os Pedidos
        Call Chama_Tela("PedidoVendaLista", colSelecao, objPedidoVenda, objEventoCodigoPV)

    Else
        'Cahama a tela que lista todas as NFiscais
        Call Chama_Tela("NF_NFFaturaLista", colSelecao, objNFiscal, objEventoCodigoNF)

    End If

End Sub

Private Sub NotaFiscal_Click()
    
Dim lErro As Long

On Error GoTo Erro_NotaFiscal_Click
    
    'Se o Foco estava com Pedido de Venda
    If iTipo_Documento <> NOTA_FISCAL And iAlterado = REGISTRO_ALTERADO Then
        
        lErro = Teste_Salva(Me, iAlterado)
        If lErro <> SUCESSO And lErro <> 20323 Then Error 58394
            
        'Caso Cancele tem que voltar o Foco para NotaFiscal
        If lErro = 20323 Then Error 58395
        
        NotaFiscal.Value = True
        
        Call Limpa_Tela_Comissoes
        
        iAlterado = 0
        
    ElseIf iTipo_Documento <> NOTA_FISCAL Then
        
        'Limpa a Tela
        Call Limpa_Tela_Comissoes
        
        iAlterado = 0
    
    End If
    
    If NotaFiscal.Value = True Then
        iTipo_Documento = NOTA_FISCAL
        Serie.Visible = True
        LabelSerie.Visible = True
    
        LabelECF.Visible = False
        ECF.Visible = False
        COO.Visible = False
        LabelCOO.Visible = False
        
        Codigo.Visible = True
        LabelNumero.Visible = True
    End If
    
    Exit Sub
    
Erro_NotaFiscal_Click:

    Select Case Err

        Case 58394 'Tratado na Rotina Chamada
        
        Case 58395
            If iTipo_Documento = PEDIDO_DE_VENDA Then PedidoVenda.Value = True
            If iTipo_Documento = CUPOM_FISCAL Then OptionCupom.Value = True
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154391)

    End Select

    Exit Sub

End Sub

Private Sub NotaFiscal_GotFocus()
    
    If NotaFiscal.Value = True Then
        iTipo_Documento = NOTA_FISCAL
    End If
        
End Sub


Private Sub objEventoCodigoPV_evSelecao(obj1 As Object)

Dim lErro As Long
Dim iIndice As Integer
Dim objPedidoVenda As ClassPedidoDeVenda

On Error GoTo Erro_objEventoCodigoPV_evSelecao

    Set objPedidoVenda = obj1
    
    Call Limpa_Tela_Comissoes

    'Preenche o Codigo
    Codigo.Text = CStr(objPedidoVenda.lCodigo)
    
    'A serie
    Serie.Text = ""
    
    'Preenche o Grid de Comissões
    lErro = PedidoDeVenda_PreencheGrid
    If lErro <> SUCESSO Then Error 21352

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoCodigoPV_evSelecao:

    Select Case Err

        Case 21352

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154392)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigoNF_evSelecao(obj1 As Object)

Dim lErro As Long
Dim iIndice As Integer
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoCodigoNF_evSelecao

    Set objNFiscal = obj1
    
    Call Limpa_Tela_Comissoes
    
    'Preenche o Codigo e a Serie
    Codigo.Text = CStr(objNFiscal.lNumNotaFiscal)
    Serie.Text = objNFiscal.sSerie
    Call DateParaMasked(DataEmissao, objNFiscal.dtDataEmissao)
    
    'Preenche o Grid de Comissões
    lErro = NotaFisc_PreencheGrid
    If lErro <> SUCESSO Then Error 21353

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoCodigoNF_evSelecao:

    Select Case Err

        Case 21353
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154393)

    End Select

    Exit Sub


End Sub

'tulio 9/5/02
Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor
Dim iIndice As Integer
Dim lErro As Long
Dim bExisteVendedor As Boolean

On Error GoTo Erro_objEventoVendedor_evSelecao

    Set objVendedor = obj1
    
    'Verifica se GridComissoes foi preenchido
    If objGrid1.iLinhasExistentes > 0 Then

        'Loop no GridComissoes
        For iIndice = 1 To objGrid1.iLinhasExistentes
        
            'Se o vendedor comparece em outra linha
            If iIndice <> GridComissoes.Row And UCase(GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col)) = UCase(Vendedor.Text) Then



        
                'intutiliza o trecho abaixo...
                bExisteVendedor = True
                                    
                'se ja tinha achado o vendedor antes
                If bExisteVendedor = True Then
                
                    'erro, pois provavelmente ja existe vendedor como direto e indireto
                    Error 25691
                
                Else
                
                    'achou vendedor
                    bExisteVendedor = True
        
                    'senao, o campo direto/indireto da linha atual recebe o q nao esta preenchido no campo da linha iindice
                    If GridComissoes.TextMatrix(iIndice, iGrid_DiretoIndireto_Col) = DiretoIndireto.List(DiretoIndireto.ItemData(VENDEDOR_DIRETO)) Then
                        
                        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_DiretoIndireto_Col) = DiretoIndireto.List(DiretoIndireto.ItemData(VENDEDOR_INDIRETO))
                        
                    Else
                        
                        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_DiretoIndireto_Col) = DiretoIndireto.List(DiretoIndireto.ItemData(VENDEDOR_DIRETO))
            
                    End If
    
                End If
                        
            End If
            
        Next

    End If
        
    'se nao encontrou o vendedor em outra linha e o campo direto/indireto esta em branco, seta o campo como direto
    If bExisteVendedor = False And Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_DiretoIndireto_Col))) = 0 Then GridComissoes.TextMatrix(GridComissoes.Row, iGrid_DiretoIndireto_Col) = DiretoIndireto.List(DiretoIndireto.ItemData(VENDEDOR_DIRETO))
    
    'Preenche o Vendedor
    Vendedor.Text = objVendedor.sNomeReduzido

    If GridComissoes.Row > 0 Then
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Vendedor_Col) = Vendedor.Text
    Else
        GridComissoes.TextMatrix(1, iGrid_Vendedor_Col) = Vendedor.Text
    End If

    Me.Show

    Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case Err

        Case 58312
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_JA_EXISTENTE", Err, objVendedor.sNomeReduzido)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154394)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim objSerie As New ClassSerie
Dim colSerie As New colSerie
Dim sNomeReduzidoVendedor As String
Dim sNomeReduzidoCliente As String

On Error GoTo Erro_Form_Load
    
    'Inicializa o Documento
    iTipo_Documento = PEDIDO_DE_VENDA
    
    sNomeReduzidoVendedor = String(STRING_VENDEDOR_NOME_REDUZIDO, 0)

    'Lê Series da tabela Serie e devolve na coleção
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then Error 21358

    For Each objSerie In colSerie
        Serie.AddItem objSerie.sSerie
    Next
    
    'Inicializa o Grid de Comissões
    lErro = Inicializa_GridComissoes()
    If lErro <> SUCESSO Then Error 21359
    
    'Inicializa os eventos de Browser
    Set objEventoVendedor = New AdmEvento
    Set objEventoCodigoPV = New AdmEvento
    Set objEventoCodigoNF = New AdmEvento
    Set objEventoCupomFiscal = New AdmEvento

    PedidoVenda.Value = True
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 21358, 21359

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154395)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Inicializa_GridComissoes() As Long

Dim iIndice As Integer

    Set objGrid1 = New AdmGrid

    'tela em questão
    Set objGrid1.objForm = Me

    'titulos do grid
    objGrid1.colColuna.Add ("")
    objGrid1.colColuna.Add ("Vendedor")
    objGrid1.colColuna.Add ("Percentual")
    objGrid1.colColuna.Add ("Valor Base")
    objGrid1.colColuna.Add ("Valor")
    objGrid1.colColuna.Add ("% Emissão")
    objGrid1.colColuna.Add ("Valor Emissão")
    objGrid1.colColuna.Add ("% Baixa")
    objGrid1.colColuna.Add ("Valor Baixa")
    objGrid1.colColuna.Add ("Direta/Indireta")

   'campos de edição do grid
    objGrid1.colCampo.Add (Vendedor.Name)
    objGrid1.colCampo.Add (PercentualComissao.Name)
    objGrid1.colCampo.Add (ValorBase.Name)
    objGrid1.colCampo.Add (ValorComissao.Name)
    objGrid1.colCampo.Add (PercentualEmissao.Name)
    objGrid1.colCampo.Add (ValorEmissao.Name)
    objGrid1.colCampo.Add (PercentualBaixa.Name)
    objGrid1.colCampo.Add (ValorBaixa.Name)
    objGrid1.colCampo.Add (DiretoIndireto.Name)
    
    'Colunas do Grid
    iGrid_Vendedor_Col = 1
    iGrid_Percentual_Comissao_Col = 2
    iGrid_Valor_Base_Col = 3
    iGrid_Valor_Comissao_Col = 4
    iGrid_Percentual_Emissao_Col = 5
    iGrid_Valor_Emissao_Col = 6
    iGrid_Percentual_Baixa_Col = 7
    iGrid_Valor_Baixa_Col = 8
    iGrid_DiretoIndireto_Col = 9

    objGrid1.objGrid = GridComissoes
    
'mario
'    Set objGridComissoes = GridComissoes

    'tulio 9/5/02
    objGrid1.objGrid.Rows = NUM_MAX_COMISSOES
    
    'linhas visiveis do grid sem contar com as linhas fixas
    objGrid1.iLinhasVisiveis = 4

    objGrid1.objGrid.ColWidth(0) = 300

    objGrid1.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid1.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGrid1)
    
    Inicializa_GridComissoes = SUCESSO

End Function

'tulio 9/5/02
Function Preenche_GridComissoes_PedidoDeVenda(colComissoes As Collection) As Long
'preenche o grid com os dados contidos na coleção colComissaoPV

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim objComissaoPV As New ClassComissaoPedVendas
Dim dTotalPercentual As Double
Dim dTotalValorComissao As Double
Dim objVendedor As New ClassVendedor
Dim sNomeReduzidoVendedor As String

On Error GoTo Erro_Preenche_GridComissoes_PedidoDeVenda

    sNomeReduzidoVendedor = String(STRING_VENDEDOR_NOME_REDUZIDO, 0)
    
    'Limpa o Grid
    GridComissoes.Clear
    
    'Inicializa o Grid
    lErro = Inicializa_GridComissoes()
    If lErro <> SUCESSO Then Error 21411

    objGrid1.iLinhasExistentes = colComissoes.Count

    dTotalPercentual = 0
    dTotalValorComissao = 0

    'preenche o grid com os dados retornados na coleção colComissoes
    For iIndice = 1 To colComissoes.Count

        Set objComissaoPV = colComissoes.Item(iIndice)
        
        objVendedor.iCodigo = objComissaoPV.iCodVendedor
        
        'Lê o código do Vendedor
        lErro = CF("Vendedor_Le", objVendedor)
        If lErro <> SUCESSO And lErro <> 12582 Then Error 64134

        'Se não encontrou --> ERRO
        If lErro = 12582 Then Error 21412
        
        'Preenche o Grid
        GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col) = objVendedor.sNomeReduzido
        GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Comissao_Col) = Format(objComissaoPV.dPercentual, "Percent")
        GridComissoes.TextMatrix(iIndice, iGrid_Valor_Base_Col) = Format(objComissaoPV.dValorBase, "Standard")
        GridComissoes.TextMatrix(iIndice, iGrid_Valor_Comissao_Col) = Format(objComissaoPV.dValor, "Standard")
        GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Emissao_Col) = Format(objComissaoPV.dPercentualEmissao, "Percent")
        GridComissoes.TextMatrix(iIndice, iGrid_Valor_Emissao_Col) = Format(objComissaoPV.dValorEmissao, "Standard")
        GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Baixa_Col) = Format(1 - objComissaoPV.dPercentualEmissao, "Percent")
        GridComissoes.TextMatrix(iIndice, iGrid_Valor_Baixa_Col) = Format(objComissaoPV.dValor - objComissaoPV.dValorEmissao, "Standard")
        
        'se vendedor for indireto
        If objComissaoPV.iIndireta = VENDEDOR_INDIRETO Then
        
            'preenche com o conteudo indicado pelo itemdata do vendedor_indireto na combo
            GridComissoes.TextMatrix(iIndice, iGrid_DiretoIndireto_Col) = DiretoIndireto.List(DiretoIndireto.ItemData(VENDEDOR_INDIRETO))
        
        'senao
        Else
            
            'preenche com o conteudo indicado pelo itemdata do vendedor_direto na combo
            GridComissoes.TextMatrix(iIndice, iGrid_DiretoIndireto_Col) = DiretoIndireto.List(DiretoIndireto.ItemData(VENDEDOR_DIRETO))
        
        End If
        
        'Calcula os totais
        dTotalPercentual = dTotalPercentual + objComissaoPV.dPercentual
        dTotalValorComissao = dTotalValorComissao + objComissaoPV.dValor

    Next
    
    'Preenche os Totais
    TotalPercentualComissao.Caption = Format(dTotalPercentual, "Percent")
    TotalValorComissao.Caption = Format(dTotalValorComissao, "Standard")

    iAlterado = 0

    Preenche_GridComissoes_PedidoDeVenda = SUCESSO

    Exit Function

Erro_Preenche_GridComissoes_PedidoDeVenda:

    Preenche_GridComissoes_PedidoDeVenda = Err

    Select Case Err

        Case 21411, 64134

        Case 21412
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", Err, objComissaoPV.iCodVendedor)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154396)

    End Select

    Exit Function

End Function

Function Preenche_GridComissoes(colComissoes As Collection) As Long
'preenche o grid com os dados contidos na coleção colComissaoPV

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
'Dim objComissaoNF As New ClassComissaoNF -> desperdicio de espaco...
Dim objComissao As Object
Dim dTotalPercentual As Double
Dim dTotalValorComissao As Double
Dim objVendedor As New ClassVendedor
Dim sNomeReduzidoVendedor As String, sDiretoIndireto As String

On Error GoTo Erro_Preenche_GridComissoes

    sNomeReduzidoVendedor = String(STRING_VENDEDOR_NOME_REDUZIDO, 0)
    
    'limpa o Grid
    GridComissoes.Clear
    
    'Atualiza as Linha Visiveis Grid
    If colComissoes.Count < objGrid1.iLinhasVisiveis Then
        objGrid1.objGrid.Rows = objGrid1.iLinhasVisiveis + 1
    Else
        objGrid1.objGrid.Rows = colComissoes.Count + 1
    End If
    
    'inicializa o Grid
    lErro = Inicializa_GridComissoes()
    If lErro <> SUCESSO Then Error 21413

    objGrid1.iLinhasExistentes = colComissoes.Count

    dTotalPercentual = 0
    dTotalValorComissao = 0

    'preenche o grid com os dados retornados na coleção colComissoes
    For iIndice = 1 To colComissoes.Count

        Set objComissao = colComissoes.Item(iIndice)

        objVendedor.iCodigo = objComissao.iCodVendedor
        
        'Lê o código do Vendedor
        lErro = CF("Vendedor_Le", objVendedor)
        If lErro <> SUCESSO And lErro <> 12582 Then Error 64134

        'Se não encontrou --> ERRO
        If lErro = 12582 Then Error 21414
   
        'Preenche o Grid
        GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col) = objVendedor.sNomeReduzido
        GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Comissao_Col) = Format(objComissao.dPercentual, "Percent")
        GridComissoes.TextMatrix(iIndice, iGrid_Valor_Base_Col) = Format(objComissao.dValorBase, "Standard")
        GridComissoes.TextMatrix(iIndice, iGrid_Valor_Comissao_Col) = Format(objComissao.dValor, "Standard")
        GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Emissao_Col) = Format(objComissao.dPercentualEmissao, "Percent")
        GridComissoes.TextMatrix(iIndice, iGrid_Valor_Emissao_Col) = Format(objComissao.dValorEmissao, "Standard")
        GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Baixa_Col) = Format(1 - objComissao.dPercentualEmissao, "Percent")
        GridComissoes.TextMatrix(iIndice, iGrid_Valor_Baixa_Col) = Format((objComissao.dValor - objComissao.dValorEmissao), "Standard")
        
        '*** 19/06/02 - INÍCIO Luiz G.F.Nogueira ***
        'Descobre o valor que será exibido no campo DiretoIndireto
        Select Case objComissao.iIndireta
        
            'Se a comissão foi para um vendedor direto
            Case VENDEDOR_DIRETO
            
                'Guarda o texto "Direto"
                sDiretoIndireto = VENDEDOR_DIRETO_STRING
            
            'Se a comissão for para um vendedor indireto
            Case VENDEDOR_INDIRETO

                'Guarda o texto "Indireto"
                sDiretoIndireto = VENDEDOR_INDIRETO_STRING
            
        End Select
        
        'preenche com o conteudo indicado pelo itemdata do vendedor_direto na combo
        GridComissoes.TextMatrix(iIndice, iGrid_DiretoIndireto_Col) = sDiretoIndireto
        '*** 19/06/02 - FIM Luiz G.F.Nogueira ***
        
        'Calcula os totais
        dTotalPercentual = dTotalPercentual + objComissao.dPercentual
        dTotalValorComissao = dTotalValorComissao + objComissao.dValor

    Next
    
    'Preenche os totais
    TotalPercentualComissao.Caption = Format(dTotalPercentual, "Percent")
    TotalValorComissao.Caption = Format(dTotalValorComissao, "Standard")

    iAlterado = 0

    Preenche_GridComissoes = SUCESSO

    Exit Function

Erro_Preenche_GridComissoes:

    Preenche_GridComissoes = Err

    Select Case Err

        Case 21413

        Case 21414
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", Err, objComissao.iCodVendedor)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154397)

    End Select

    Exit Function

End Function

Function Preenche_GridComissoes_CF(colComissoes As Collection) As Long
'preenche o grid com os dados contidos na coleção colComissoes

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim objComissao As ClassComissoesCF
Dim dTotalPercentual As Double
Dim dTotalValorComissao As Double
Dim objVendedor As New ClassVendedor
Dim sNomeReduzidoVendedor As String, sDiretoIndireto As String
Dim dTotalValorBase As Double

On Error GoTo Erro_Preenche_GridComissoes_CF

    sNomeReduzidoVendedor = String(STRING_VENDEDOR_NOME_REDUZIDO, 0)
    
    'limpa o Grid
    GridComissoes.Clear
    
    'Atualiza as Linha Visiveis Grid
    If colComissoes.Count < objGrid1.iLinhasVisiveis Then
        objGrid1.objGrid.Rows = objGrid1.iLinhasVisiveis + 1
    Else
        objGrid1.objGrid.Rows = colComissoes.Count + 1
    End If
    
    'inicializa o Grid
    lErro = Inicializa_GridComissoes()
    If lErro <> SUCESSO Then gError 126325

    objGrid1.iLinhasExistentes = colComissoes.Count

    dTotalPercentual = 0
    dTotalValorComissao = 0

    'preenche o grid com os dados retornados na coleção colComissoes
    For iIndice = 1 To colComissoes.Count

        Set objComissao = colComissoes.Item(iIndice)

        objVendedor.iCodigo = objComissao.iCodVendedor
        
        'Lê o código do Vendedor
        lErro = CF("Vendedor_Le", objVendedor)
        If lErro <> SUCESSO And lErro <> 12582 Then gError 126326

        'Se não encontrou --> ERRO
        If lErro = 12582 Then gError 126327
   
        'Preenche o Grid
        GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col) = objVendedor.sNomeReduzido
        GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Comissao_Col) = Format(objComissao.dValorComissao / objComissao.dValorBase, "Percent")
        GridComissoes.TextMatrix(iIndice, iGrid_Valor_Base_Col) = Format(objComissao.dValorBase, "Standard")
        GridComissoes.TextMatrix(iIndice, iGrid_Valor_Comissao_Col) = Format(objComissao.dValorComissao, "Standard")
        GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Emissao_Col) = Format(1, "Percent")
        GridComissoes.TextMatrix(iIndice, iGrid_Valor_Emissao_Col) = Format(objComissao.dValorComissao, "Standard")
        GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Baixa_Col) = Format(0, "Percent")
        GridComissoes.TextMatrix(iIndice, iGrid_Valor_Baixa_Col) = Format(0, "Standard")
        
        'Descobre o valor que será exibido no campo DiretoIndireto
        Select Case objComissao.iIndireta
        
            'Se a comissão foi para um vendedor direto
            Case VENDEDOR_DIRETO
            
                'Guarda o texto "Direto"
                sDiretoIndireto = VENDEDOR_DIRETO_STRING
            
            'Se a comissão for para um vendedor indireto
            Case VENDEDOR_INDIRETO

                'Guarda o texto "Indireto"
                sDiretoIndireto = VENDEDOR_INDIRETO_STRING
            
        End Select
        
        'preenche com o conteudo indicado pelo itemdata do vendedor_direto na combo
        GridComissoes.TextMatrix(iIndice, iGrid_DiretoIndireto_Col) = sDiretoIndireto
        
        'Calcula os totais
        dTotalValorBase = dTotalValorBase + objComissao.dValorBase
        dTotalValorComissao = dTotalValorComissao + objComissao.dValorComissao

    Next
    
    If dTotalValorBase > 0 Then
    
        'Preenche os totais
        TotalPercentualComissao.Caption = Format(dTotalValorComissao / dTotalValorBase, "Percent")
        TotalValorComissao.Caption = Format(dTotalValorComissao, "Standard")
        TotalValorBase.Caption = Format(dTotalValorBase, "Standard")

    End If

    iAlterado = 0

    Preenche_GridComissoes_CF = SUCESSO

    Exit Function

Erro_Preenche_GridComissoes_CF:

    Preenche_GridComissoes_CF = gErr

    Select Case gErr

        Case 126325, 126326

        Case 126327
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", gErr, objComissao.iCodVendedor)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154398)

    End Select

    Exit Function

End Function


Function PedidoDeVenda_PreencheGrid() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objPedidoDeVenda As New ClassPedidoDeVenda
Dim sNomeReduzidoCliente As String
Dim objcliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_PedidoDeVenda_PreencheGrid

    objPedidoDeVenda.lCodigo = CLng(Codigo.Text)
    objPedidoDeVenda.iFilialEmpresa = giFilialEmpresa

    'Lê o Pedido de Venda
    lErro = CF("PedidoDeVenda_Le", objPedidoDeVenda)
    If lErro <> SUCESSO And lErro <> 26509 Then Error 16951

    'Não achou o Pedido de Venda
    If lErro = 26509 Then Error 16952

    objcliente.lCodigo = objPedidoDeVenda.lCliente
    'Lê o Cliente
    lErro = CF("Cliente_Le", objcliente)
    If lErro <> SUCESSO And lErro <> 12293 Then Error 19359

    'Se não achou o Cliente --> Erro
    If lErro <> SUCESSO Then Error 21417

    Cliente.Caption = objcliente.sNomeReduzido
    Valor.Caption = Format(objPedidoDeVenda.dValorTotal, "Standard")
    Call DateParaMasked(DataEmissao, objPedidoDeVenda.dtDataEmissao)

    objFilialCliente.lCodCliente = objPedidoDeVenda.lCliente
    objFilialCliente.iCodFilial = objPedidoDeVenda.iFilial
    'Lê a Filial do Cliente
    lErro = CF("FilialCliente_Le", objFilialCliente)
    If lErro <> SUCESSO And lErro <> 12567 Then Error 43671

    'Se não achou a Filial do Cliente --> Erro
    If lErro <> SUCESSO Then Error 43672

    Filial.Caption = CStr(objPedidoDeVenda.iFilial) & SEPARADOR & objFilialCliente.sNome

    lErro = CF("ComissoesPV_Le", objPedidoDeVenda)
    If lErro <> SUCESSO And lErro <> 21367 Then Error 21415

    lErro = Preenche_GridComissoes_PedidoDeVenda(objPedidoDeVenda.colComissoes)
    If lErro <> SUCESSO Then Error 21418

    PedidoDeVenda_PreencheGrid = SUCESSO

    Exit Function

Erro_PedidoDeVenda_PreencheGrid:

    PedidoDeVenda_PreencheGrid = Err

    Select Case Err

        Case 16952
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", Err, objPedidoDeVenda.lCodigo)
            Codigo.SetFocus

        Case 21415, 16951, 19359, 21418, 43671

        Case 43672
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", Err, objPedidoDeVenda.iFilial, objPedidoDeVenda.lCliente)

        Case 21417
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objcliente.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154399)

    End Select

    Exit Function

End Function

Function NotaFisc_PreencheGrid() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objNFiscal As New ClassNFiscal
Dim objcliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_NotaFisc_PreencheGrid

    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 89999

    objNFiscal.lNumNotaFiscal = CLng(Codigo.Text)
    objNFiscal.sSerie = LTrim(Serie.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.dtDataEmissao = CDate(DataEmissao.Text)

    'TEM QUE LER A PROPRIA NF
    lErro = CF("NFiscal_Le_NumeroSerie", objNFiscal)
    If lErro <> SUCESSO And lErro <> 43676 Then gError 43677

    'Se não encontrou a Nota Fiscal --> Erro
    If lErro <> SUCESSO Then gError 43678

    objcliente.lCodigo = objNFiscal.lCliente
    
    'Lê o Cliente
    lErro = CF("Cliente_Le", objcliente)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 19357

    'Se não achou o Cliente --> Erro
    If lErro <> SUCESSO Then gError 21446

    Cliente.Caption = objcliente.sNomeReduzido
    Valor.Caption = Format(objNFiscal.dValorTotal, "Standard")

    objFilialCliente.lCodCliente = objNFiscal.lCliente
    objFilialCliente.iCodFilial = objNFiscal.iFilialCli
    
    'Lê a Filial do Cliente
    lErro = CF("FilialCliente_Le", objFilialCliente)
    If lErro <> SUCESSO And lErro <> 12567 Then gError 43667

    'Se não achou a Filial do Cliente --> Erro
    If lErro <> SUCESSO Then gError 43668

    Filial.Caption = CStr(objNFiscal.iFilialCli) & SEPARADOR & objFilialCliente.sNome

    'Lê as comissões
    lErro = CF("NFiscal_Le_Comissoes", objNFiscal)
    If lErro <> SUCESSO And lErro <> 21386 Then gError 21419

    lErro = Preenche_GridComissoes(objNFiscal.ColComissoesNF)
    If lErro <> SUCESSO Then gError 21422

    NotaFisc_PreencheGrid = SUCESSO

    Exit Function

Erro_NotaFisc_PreencheGrid:

    NotaFisc_PreencheGrid = gErr

    Select Case gErr

        Case 19357, 21419, 21422, 43667, 43677

        Case 21446
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objcliente.lCodigo)

        Case 43668
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", gErr, objNFiscal.iFilialCli, objNFiscal.lCliente)

        Case 43678
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", gErr, objNFiscal.lNumNotaFiscal)

        Case 89999
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154400)

    End Select

    Exit Function

End Function

Function CupomFiscal_PreencheGrid() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sNomeReduzidoCliente As String
Dim objcliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim objCupomFiscal As New ClassCupomFiscal
Dim iCodFilial As Integer
Dim colComissao As New Collection

On Error GoTo Erro_CupomFiscal_PreencheGrid

    objCupomFiscal.lNumero = StrParaLong(COO.Text)
    objCupomFiscal.iECF = StrParaInt(ECF.Text)
    objCupomFiscal.iFilialEmpresa = giFilialEmpresa

    'Lê o Cupom Fiscal
    lErro = CF("CupomFiscal_Le", objCupomFiscal)
    If lErro <> SUCESSO And lErro <> 105262 Then gError 126313

    'Não achou o Cupom Fiscal
    If lErro = 105262 Then gError 126314

    Valor.Caption = Format(objCupomFiscal.dValorTotal, "Standard")
    Call DateParaMasked(DataEmissao, objCupomFiscal.dtDataEmissao)

    If Len(objCupomFiscal.sCPFCGC) > 0 Then

        objcliente.sCgc = objCupomFiscal.sCPFCGC
        
        'Lê o Cliente
        lErro = CF("Cliente_Le_CGC", objcliente, iCodFilial)
        If lErro <> SUCESSO And lErro <> 6710 Then gError 126315
    
        'Se achou o Cliente
        If lErro = SUCESSO Then
    
            Cliente.Caption = objcliente.sNomeReduzido
        
            objFilialCliente.lCodCliente = objcliente.lCodigo
            objFilialCliente.iCodFilial = iCodFilial
        
            'Lê a Filial do Cliente
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then gError 126317
    
            'Se achou a Filial do Cliente
            If lErro = SUCESSO Then Filial.Caption = CStr(objFilialCliente.iCodFilial) & SEPARADOR & objFilialCliente.sNome
                
        End If

    End If

    lErro = CF("ComissoesCF_Le", objCupomFiscal, colComissao)
    If lErro <> SUCESSO Then gError 126322

    lErro = Preenche_GridComissoes_CF(colComissao)
    If lErro <> SUCESSO Then Error 126323

    CupomFiscal_PreencheGrid = SUCESSO

    Exit Function

Erro_CupomFiscal_PreencheGrid:

    CupomFiscal_PreencheGrid = gErr

    Select Case gErr

        Case 126313, 126315, 126317, 126322, 126323

        Case 126314
            Call Rotina_Erro(vbOKOnly, "ERRO_CUPOM_FISCAL_NAO_CADASTRADO2", gErr, objCupomFiscal.lNumero, objCupomFiscal.iFilialEmpresa, objCupomFiscal.iECF)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154401)

    End Select

    Exit Function

End Function

Private Sub GridComissoes_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridComissoes_GotFocus()
    
    Call Grid_Recebe_Foco(objGrid1)
    
End Sub

Private Sub GridComissoes_EnterCell()
    
    Call Grid_Entrada_Celula(objGrid1, iAlterado)
    
End Sub

Private Sub GridComissoes_LeaveCell()
    
    Call Saida_Celula(objGrid1)
    
End Sub

Private Sub GridComissoes_KeyDown(KeyCode As Integer, Shift As Integer)
    
Dim iIndice As Integer
Dim dTotalValor As Double
Dim dTotalPercentual As Double

    Call Grid_Trata_Tecla1(KeyCode, objGrid1)
    
    dTotalValor = 0
    dTotalPercentual = 0

    For iIndice = 1 To objGrid1.iLinhasExistentes

        If Len(GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Comissao_Col)) > 0 Then dTotalPercentual = dTotalPercentual + CDbl(left(GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Comissao_Col), Len(GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Comissao_Col)) - 1))

        If Len(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Comissao_Col)) > 0 Then dTotalValor = dTotalValor + CDbl(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Comissao_Col))

    Next

    TotalPercentualComissao.Caption = Format(dTotalPercentual, "Fixed") & "%"
    TotalValorComissao.Caption = Format(dTotalValor, "Standard")
    
End Sub

Private Sub GridComissoes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridComissoes_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGrid1)
    
End Sub

Private Sub GridComissoes_RowColChange()
    Call Grid_RowColChange(objGrid1)
End Sub

Private Sub GridComissoes_Scroll()
    Call Grid_Scroll(objGrid1)
End Sub

'tulio 9/5/02
Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da celula do grid que está deixando de ser a corrente /m

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case GridComissoes.Col

            Case iGrid_Vendedor_Col
                
                'Faz a Critica para o Vendedor
                lErro = Saida_Celula_Vendedor(objGridInt)
                If lErro <> SUCESSO Then Error 21432

            Case iGrid_Valor_Base_Col
                
                'Faz a critica para o Valor Base
                lErro = Saida_Celula_ValorBase(objGridInt)
                If lErro <> SUCESSO Then Error 21424

            Case iGrid_Valor_Emissao_Col
                
                'Faz a critica para o Valor Emissao
                lErro = Saida_Celula_ValorEmissao(objGridInt)
                If lErro <> SUCESSO Then Error 21430

            Case iGrid_Percentual_Comissao_Col
            
                'Faz a critica para o Percentual Comissão
                lErro = Saida_Celula_PercentualComissao(objGridInt)
                If lErro <> SUCESSO Then Error 21431

            Case iGrid_Percentual_Emissao_Col
                
                'Faz a critica para o Percentual Emissão
                lErro = Saida_Celula_PercentualEmissao(objGridInt)
                If lErro <> SUCESSO Then Error 21425
            
            Case iGrid_Valor_Comissao_Col
                
                'Faz a critica para o Valor Comissao
                lErro = Saida_Celula_ValorComissao(objGridInt)
                If lErro <> SUCESSO Then Error 58307
                
            Case iGrid_DiretoIndireto_Col
            
                'Faz a critica para o Valor Comissao
                lErro = Saida_Celula_DiretoIndireto(objGridInt)
                If lErro <> SUCESSO Then Error 58308
                
        End Select

        iAlterado = 1

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 21426

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 21424, 21425, 21430, 21431, 21432, 21426, 58307, 58308

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154402)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_ValorEmissao(objGridInt As AdmGrid) As Long
'Faz a crítica da celula Valor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentual As Double
Dim dValorComissao As Double
Dim dValorEmissao As Double
Dim lComprimento As Long

On Error GoTo Erro_Saida_Celula_ValorEmissao

    Set objGridInt.objControle = ValorEmissao

    'Verifica se valor está preenchido
    If Len(ValorEmissao.ClipText) > 0 Then

        'Critica se valor base é positivo
        lErro = Valor_NaoNegativo_Critica(ValorEmissao.Text)
        If lErro <> SUCESSO Then Error 21439

        dValorEmissao = CDbl(ValorEmissao.Text)

        'Mostra na tela o Valor
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Emissao_Col) = Format(dValorEmissao, "Fixed")

        'Verifica se valor Comissao correspondente está preenchido
        If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Comissao_Col))) > 0 Then

            dValorComissao = CDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Comissao_Col))

            If dValorEmissao > dValorComissao Then Error 21450

            lComprimento = Len(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Emissao_Col))

            'Verifica se percentual emissao correspondente está preenchido
            If lComprimento > 0 Then dPercentual = PercentParaDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Emissao_Col))

            If (dPercentual * dValorComissao) <> dValorEmissao Then

                dPercentual = dValorEmissao / dValorComissao

                'Mostra o percentual da comissao na tela
                GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Emissao_Col) = Format(dPercentual, "Percent")

                'Coloca o percentual na baixa na tela
                GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Baixa_Col) = Format(1 - dPercentual, "Percent")

                'Coloca o valor na baixa na tela
                If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Comissao_Col))) > 0 Then
                    ValorBaixa.Text = Format(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Comissao_Col) - (GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Emissao_Col)), "Standard")
                    GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Baixa_Col) = Format(ValorBaixa.Text, "Standard")
                End If
                
            End If

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridComissoes.Row - GridComissoes.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 21440

    Saida_Celula_ValorEmissao = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorEmissao:

    Saida_Celula_ValorEmissao = Err

    Select Case Err

        Case 21439, 21440
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 21450
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_EMISSAO_MAIOR", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            ValorEmissao.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154403)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_ValorBase(objGridInt As AdmGrid) As Long
'Faz a crítica da celula ValorBase do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentual As Double
Dim dValorBase As Double
Dim dValorComissao As Double
Dim dValorEmissao As Double
Dim dValorBaixa As Double
Dim dTotalPercentual As Double
Dim dTotalValor As Double
Dim dValorDoc As Double
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_ValorBase

    Set objGridInt.objControle = ValorBase

    'Verifica se valor base está preenchido
    If Len(ValorBase.ClipText) > 0 Then

        'Critica se valor base é positivo
        lErro = Valor_Positivo_Critica(ValorBase.Text)
        If lErro <> SUCESSO Then Error 21437

        dValorBase = CDbl(ValorBase.Text)

        'Mostra na tela o ValorBase formatado
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Base_Col) = Format(dValorBase, "Standard")
        
        'Verifica se o valor do documento está preenchido
        If Len(Trim(Valor.Caption)) <> 0 Then
        
            dValorDoc = CDbl(Valor.Caption)
            
            'Se Valor Base é maior que o Valor do Documento --> Erro
            If dValorBase > dValorDoc Then Error 17197
            
        End If
        
        'Verifica se percentual comissao está preenchido
        If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Comissao_Col))) > 0 Then

            dPercentual = PercentParaDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Comissao_Col))

            'Calcula o valor da comissao
            dValorComissao = dPercentual * dValorBase

            'Mostra na tela o valor da comissao
            GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Comissao_Col) = Format(dValorComissao, "Standard")
            
            'Se o Percentual de Emissao estiver preenchido
            If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Emissao_Col))) > 0 Then

                dPercentual = PercentParaDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Emissao_Col))
                
                'Calcula o valor de Emissão
                dValorEmissao = dPercentual * dValorComissao
                
                'Preenche o valor de Emissão
                GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Emissao_Col) = Format(dValorEmissao, "Standard")
                
                'Calcula o Velor da Baixa
                dValorBaixa = dValorComissao - dValorEmissao
                
                'Preenche o valor da Baixa
                GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Baixa_Col) = Format(CStr(dValorBaixa), "Standard")
                
                'Preenche o Percentual da Baixa
                GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Baixa_Col) = Format(1 - dPercentual, "Percent")

            Else
                
                'Se o Valor da Emissao estiver Preenchido
                If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Emissao_Col))) > 0 Then
                    
                    dValorEmissao = CDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Emissao_Col))
                    
                    'Se o Valor da Emissão for maior que o da Base --> EERO
                    If dValorEmissao > dValorBase Then Error 21452
                    
                    'Calcula o Percentual da Emissao
                    dPercentual = dValorEmissao / dValorComissao
                    
                    'Preenche o Percentual da Emissão
                    GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Emissao_Col) = Format(dPercentual, "Percent")
                    
                    'Calcula o Valor da Baixa
                    dValorBaixa = dValorComissao - dValorEmissao
                    
                    'Preenche o Valor da Baixa
                    GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Baixa_Col) = Format(CStr(dValorBaixa), "Standard")
                    
                    'Preenche o Percentual da Baixa
                    GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Baixa_Col) = Format(1 - dPercentual, "Percent")

                End If

            End If

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridComissoes.Row - GridComissoes.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    'Calcula os totais
    dTotalValor = 0
    dTotalPercentual = 0

    For iIndice = 1 To objGridInt.iLinhasExistentes
        If Len(GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Comissao_Col)) > 0 Then dTotalPercentual = dTotalPercentual + CDbl(left(GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Comissao_Col), Len(GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Comissao_Col)) - 1))

        If Len(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Comissao_Col)) > 0 Then dTotalValor = dTotalValor + CDbl(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Comissao_Col))

    Next
    
    'Preenche os totais
    TotalPercentualComissao.Caption = Format(dTotalPercentual, "Fixed") & "%"
    TotalValorComissao.Caption = Format(dTotalValor, "Standard")

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 21438

    Saida_Celula_ValorBase = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorBase:

    Saida_Celula_ValorBase = Err

    Select Case Err

        Case 21437, 21438
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 21452
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_EMISSAO_MAIOR", Err, dValorEmissao, dValorComissao)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 17197
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORBASE_MAIOR_VALORDOC", Err, dValorBase, dValorDoc)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154404)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PercentualComissao(objGridInt As AdmGrid) As Long
'Faz a crítica da celula PercentualComissoes do grid que está deixando de ser o corrente

Dim lErro As Long
Dim dPercentual As Double
Dim dValorBase As Double
Dim dValorBaixa As Double
Dim dValorEmissao As Double
Dim dValorComissao As Double
Dim dTotalPercentual As Double
Dim dTotalValor As Double
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_PercentualComissao

    Set objGridInt.objControle = PercentualComissao

    'Verifica se o percentual está preenchido
    If Len(PercentualComissao.ClipText) > 0 Then

        'Critica se é porcentagem
        lErro = Porcentagem_Critica(PercentualComissao.Text)
        If lErro <> SUCESSO Then Error 21435

        dPercentual = CDbl(PercentualComissao.Text)

        'Mostra na tela o percentual formatado
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Comissao_Col) = Format(dPercentual, "Fixed")

        'Verifica se valorbase correspondente esta preenchido
        If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Base_Col))) > 0 Then

            dValorBase = CDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Base_Col))

            'Calcula o valorcomissao
            dValorComissao = dPercentual * dValorBase / 100

            'Coloca o valorcomissoes na tela
            GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Comissao_Col) = Format(dValorComissao, "Standard")
            
            'Se o Percentual de Emissão estiver Preencheido
            If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Emissao_Col))) > 0 Then

                dPercentual = CDbl(left(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Emissao_Col), Len(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Emissao_Col)) - 1))
                
                'Calcula o Valor da Emissão
                dValorEmissao = dPercentual * dValorComissao / 100
                
                'Preeenche o Valor da Emissão
                GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Emissao_Col) = Format(CStr(dValorEmissao), "Standard")
                
                'Calcula o valor da Baixa
                dValorBaixa = dValorComissao - dValorEmissao
                
                'Preenche o valor da Baixa
                GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Baixa_Col) = Format(CStr(dValorBaixa), "Standard")
                
                'preecnhe o Percentual da Baixa
                GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Baixa_Col) = Format(1 - dPercentual / 100, "Percent")


            Else
                
                'Se o valor estiver Preenchido
                If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Emissao_Col))) > 0 Then
                
                    dValorEmissao = CDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Emissao_Col))
                    
                    'Se o Valor da Emissão for maior que Valor da Comissao --> erro
                    If dValorEmissao > dValorComissao Then Error 21451
                    
                    'Calcula o Percentual de Emissão
                    dPercentual = dValorEmissao / dValorComissao * 100
            
                    'Preenche o Percentual de Emissão
                    GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Emissao_Col) = Format(dPercentual, "Fixed")
                    
                    'Calcula o Valor da Baixa
                    dValorBaixa = dValorComissao - dValorEmissao
                    
                    'Preenche o Valor da Baixa
                    GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Baixa_Col) = Format(CStr(dValorBaixa), "Standard")
                    
                    'Preenche o Percentual da Baixa
                    GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Baixa_Col) = Format(1 - dPercentual, "Percent")

                End If

            End If


        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridComissoes.Row - GridComissoes.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    'Calcula Totais
    dTotalValor = 0
    dTotalPercentual = 0

    For iIndice = 1 To objGridInt.iLinhasExistentes

        If Len(GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Comissao_Col)) > 0 Then dTotalPercentual = dTotalPercentual + CDbl(left(GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Comissao_Col), Len(GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Comissao_Col)) - 1))

        If Len(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Comissao_Col)) > 0 Then dTotalValor = dTotalValor + CDbl(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Comissao_Col))

    Next
    
    'Preenche Totais
    TotalPercentualComissao.Caption = Format(dTotalPercentual, "Fixed") & "%"
    TotalValorComissao.Caption = Format(dTotalValor, "Standard")

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 21436

    Saida_Celula_PercentualComissao = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentualComissao:

    Saida_Celula_PercentualComissao = Err

    Select Case Err

        Case 21435, 21436
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 21451
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_EMISSAO_MAIOR", Err, CStr(dValorEmissao), CStr(dValorComissao))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            PercentualComissao.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154405)

    End Select

    Exit Function

End Function

'tulio 9/5/02
Private Function Saida_Celula_ValorComissao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentual As Double
Dim dValorBase As Double
Dim dValorComissao As Double
Dim lComprimento As Long
Dim iIndice As Integer
Dim dTotalValor As Double
Dim dTotalPercentual As Double

On Error GoTo Erro_Saida_Celula_ValorComissao

    Set objGridInt.objControle = ValorComissao

    'Verifica se valor está preenchido
    If Len(ValorComissao.ClipText) > 0 Then

        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(ValorComissao.Text)
        If lErro <> SUCESSO Then Error 58302

        dValorComissao = CDbl(ValorComissao.Text)

        'Mostra na tela o Valor
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Comissao_Col) = Format(dValorComissao, "Fixed")

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 58303

    'Verifica se valor base correspondente está preenchido
    If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Base_Col))) > 0 Then

        dValorBase = CDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Base_Col))

        If dValorBase < dValorComissao Then Error 58304

        'Calcula percent de comissao
        dPercentual = (dValorComissao / dValorBase)

        'Mostra o percentual da comissao na tela
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Comissao_Col) = Format(dPercentual, "Percent")

        'Mostra valores (percentuais) de comissoes na emissao / baixa
        lErro = EmissaoBaixa_Calcula(dValorComissao)
        If lErro <> SUCESSO Then Error 58305

    End If
    
    'Calcula Totais
    dTotalValor = 0
    dTotalPercentual = 0

    For iIndice = 1 To objGridInt.iLinhasExistentes
        If Len(GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Comissao_Col)) > 0 Then dTotalPercentual = dTotalPercentual + CDbl(left(GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Comissao_Col), Len(GridComissoes.TextMatrix(iIndice, iGrid_Percentual_Comissao_Col)) - 1))

        If Len(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Comissao_Col)) > 0 Then dTotalValor = dTotalValor + CDbl(GridComissoes.TextMatrix(iIndice, iGrid_Valor_Comissao_Col))

    Next
   
    'Preenche Totais
    TotalPercentualComissao.Caption = Format(dTotalPercentual, "Fixed") & "%"
    TotalValorComissao.Caption = Format(dTotalValor, "Standard")
   
    Saida_Celula_ValorComissao = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorComissao:

    Saida_Celula_ValorComissao = Err

    Select Case Err

        Case 58304
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_COMISSAO_MAIOR_VALORBASE", Err, dValorComissao, dValorBase)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 58303, 58305, 58302
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154406)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PercentualEmissao(objGridInt As AdmGrid) As Long
'Faz a crítica da celula PercentualComissoes do grid que está deixando de ser o corrente

Dim lErro As Long
Dim dPercentual As Double
Dim dValorBase As Double
Dim dValorComissao As Double

On Error GoTo Erro_Saida_Celula_PercentualEmissao

    Set objGridInt.objControle = PercentualEmissao

    'Verifica se o percentual está preenchido
    If Len(PercentualEmissao.ClipText) > 0 Then

        'Critica se é porcentagem
        lErro = Porcentagem_Critica(PercentualEmissao.Text)
        If lErro <> SUCESSO Then gError 21433

        dPercentual = CDbl(PercentualEmissao.Text)
        
        If OptionCupom.Value = True And dPercentual <> 100 Then gError 126350
        
        'Mostra na tela o percentual formatado
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Emissao_Col) = Format(dPercentual, "Fixed")

        PercentualBaixa.Text = CStr(100 - dPercentual)
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Baixa_Col) = Format(CDbl(PercentualBaixa.Text) / 100, "Percent")

        'Verifica se valorbase correspondente esta preenchido
        If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Comissao_Col))) > 0 Then

           dValorBase = CDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Comissao_Col))

           'Calcula o valorcomissao
           dValorComissao = dPercentual * dValorBase / 100

           'Coloca o valorcomissoes na tela
           GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Emissao_Col) = Format(dValorComissao, "Standard")

           ValorBaixa.Text = Format(CStr(CDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Comissao_Col)) - CDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Emissao_Col))), "Standard")
           GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Baixa_Col) = Format(ValorBaixa.Text, "Standard")

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridComissoes.Row - GridComissoes.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 21434

    Saida_Celula_PercentualEmissao = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentualEmissao:

    Saida_Celula_PercentualEmissao = gErr

    Select Case gErr

        Case 21433, 21434
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 126350
            Call Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_EMISSAO_CF", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154407)

    End Select

    Exit Function

End Function

'tulio 9/5/02
Public Function Saida_Celula_Vendedor(objGridInt As AdmGrid) As Long
'Faz a crítica da celula vendedor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim objVendedor As New ClassVendedor
Dim dValorComissao As Double
Dim vbMsgRes As VbMsgBoxResult
Dim bExisteVendedor As Boolean

On Error GoTo Erro_Saida_Celula_Vendedor

    Set objGridInt.objControle = Vendedor

    'Verifica se vendedor está preenchido
    If Len(Trim(Vendedor.Text)) > 0 Then

        'Verifica se Vendedor existe
        lErro = TP_Vendedor_Grid(Vendedor, objVendedor)
        If lErro <> SUCESSO And lErro <> 25018 And lErro <> 25020 Then Error 21441
        
        'Não achou o nome do vendedor
        If lErro = 25018 Then Error 21442
        
        'Não achou o código do vendedor
        If lErro = 25020 Then Error 21443

        'Verifica se GridComissoes foi preenchido
        If objGrid1.iLinhasExistentes > 0 Then

            'Loop no GridComissoes
            For iIndice = 1 To objGrid1.iLinhasExistentes
            
                'Se o vendedor comparece em outra linha
                If iIndice <> GridComissoes.Row And UCase(GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col)) = UCase(Vendedor.Text) Then
            
                    'intutiliza o trecho abaixo...
                    bExisteVendedor = True
                                        
                    'se ja tinha achado o vendedor antes
                    If bExisteVendedor = True Then
                    
                        'erro, pois provavelmente ja existe vendedor como direto e indireto
                        Error 25691
                    


                    Else
                    
                        'achou vendedor
                        bExisteVendedor = True
            
                        'senao, o campo direto/indireto da linha atual recebe o q nao esta preenchido no campo da linha iindice
                        If GridComissoes.TextMatrix(iIndice, iGrid_DiretoIndireto_Col) = DiretoIndireto.List(DiretoIndireto.ItemData(VENDEDOR_DIRETO)) Then
                            
                            GridComissoes.TextMatrix(GridComissoes.Row, iGrid_DiretoIndireto_Col) = DiretoIndireto.List(DiretoIndireto.ItemData(VENDEDOR_INDIRETO))
                            
                        Else
                            
                            GridComissoes.TextMatrix(GridComissoes.Row, iGrid_DiretoIndireto_Col) = DiretoIndireto.List(DiretoIndireto.ItemData(VENDEDOR_DIRETO))
                
                        End If
        
                    End If
                            
                End If
                
            Next

        End If
        
        'se nao encontrou o vendedor em outra linha e o campo direto/indireto esta em branco, seta o campo como direto
        If bExisteVendedor = False And Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_DiretoIndireto_Col))) = 0 Then GridComissoes.TextMatrix(GridComissoes.Row, iGrid_DiretoIndireto_Col) = DiretoIndireto.List(DiretoIndireto.ItemData(VENDEDOR_DIRETO))

        'se a empresa nao usa regras para o calc. de comissoes
        If gobjCRFAT.iUsaComissoesRegras = Not USA_REGRAS Then

            'Coloca perencentuais de comissão na emissão e na baixa
            GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Emissao_Col) = Format(objVendedor.dPercComissaoEmissao, "Percent")
            GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Baixa_Col) = Format(1 - objVendedor.dPercComissaoEmissao, "Percent")

        End If
        
        'Testa se valor de comissão está preenchido
        If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Comissao_Col))) > 0 Then

            dValorComissao = CDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Comissao_Col))

            'Calcula comissões na emissão e na baixa
            lErro = EmissaoBaixa_Calcula(dValorComissao)
            If lErro <> SUCESSO Then Error 22966

        End If
        
        'Acrescenta uma linha no Grid se for o caso
        If GridComissoes.Row - GridComissoes.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 21445

    Saida_Celula_Vendedor = SUCESSO

    Exit Function

Erro_Saida_Celula_Vendedor:

    Saida_Celula_Vendedor = Err

    Select Case Err

        Case 21441, 21445, 22966
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 21442 'Não encontrou nome reduzido de vendedor no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR")

            If vbMsgRes = vbYes Then

                'Preenche objVendedor com nome reduzido
                objVendedor.sNomeReduzido = Vendedor.Text

                'Chama a tela de Vendedores
                Call Chama_Tela("Vendedores", objVendedor)
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 21443 'Não encontrou codigo do vendedor no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR")

            If vbMsgRes = vbYes Then

                'Prenche objVendedor com codigo
                objVendedor.iCodigo = CDbl(Vendedor.Text)

                'Chama a tela de Vendedores
                Call Chama_Tela("Vendedores", objVendedor)
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 21444, 25691
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_JA_EXISTENTE", Err, objVendedor.sNomeReduzido)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154408)

    End Select

    Exit Function

End Function

'tulio 9/5/02
Public Function Saida_Celula_DiretoIndireto(objGridInt As AdmGrid) As Long
'Faz a crítica da celula vendedor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim bAchou As Boolean

On Error GoTo Erro_Saida_Celula_DiretoIndireto

    Set objGridInt.objControle = DiretoIndireto

    'Verifica se vendedor está preenchido
    If Len(Trim(Vendedor.Text)) > 0 Then

        
        'testa se o vendedor da linha atual ja esta em outra linha...
        For iIndice = 1 To objGrid1.iLinhasExistentes

            'Se o vendedor comparece em outra linha
            If iIndice <> GridComissoes.Row And UCase(GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col)) = UCase(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Vendedor_Col)) Then
                
                'inutiliza o trecho de codigo abaixo.. q posteriormente deve ficar habilitado...
                bAchou = True
                
                'se ja achou vendedor, erro, 3o vendedor
                'provavelmente nunca ocorrera de achar o 3o, pois o saida_celula_vendedor
                'garante isso, mas inseri esse trecho por seguranca.. tulio
                If bAchou = True Then
                
                    gError 98990

                Else
                
                    'se nao achou, marca q agora achou
                    bAchou = True
                
                    'senao, o campo direto/indireto da linha atual recebe o q nao esta preenchido no campo da linha iindice
                        If GridComissoes.TextMatrix(iIndice, iGrid_DiretoIndireto_Col) = DiretoIndireto.List(DiretoIndireto.ItemData(VENDEDOR_DIRETO)) Then
                            
                            GridComissoes.TextMatrix(GridComissoes.Row, iGrid_DiretoIndireto_Col) = DiretoIndireto.List(DiretoIndireto.ItemData(VENDEDOR_INDIRETO))
                            DiretoIndireto.ListIndex = VENDEDOR_INDIRETO
                
                        Else
                            
                            GridComissoes.TextMatrix(GridComissoes.Row, iGrid_DiretoIndireto_Col) = DiretoIndireto.List(DiretoIndireto.ItemData(VENDEDOR_DIRETO))
                            DiretoIndireto.ListIndex = VENDEDOR_DIRETO
                
                        End If
                                    
                End If
                
                'se valor de diretoindireto da linha atual eh o mesmo da outra linha...
'                If GridComissoes.TextMatrix(GridComissoes.Row, iGrid_DiretoIndireto_Col) = GridComissoes.TextMatrix(iIndice, iGrid_DiretoIndireto_Col) Then gError 98990

            End If

        Next
        
    End If
    
'    'verifica se o vendedor ja foi preenchido, senao limpa o conteudo do campo direto/indireto
'    If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_DiretoIndireto_Col))) > 0 Then
'
'        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_DiretoIndireto_Col) = STRING_VAZIO
'        DiretoIndireto.ListIndex = -1
'
'    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 21437

    Saida_Celula_DiretoIndireto = SUCESSO

    Exit Function

Erro_Saida_Celula_DiretoIndireto:

    Saida_Celula_DiretoIndireto = gErr
    
    Select Case gErr
    
        Case 21437
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 98990
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_JA_EXISTENTE2", gErr, GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 154409)
            
    End Select
    
    Exit Function

End Function

Function EmissaoBaixa_Calcula(dValorComissao As Double) As Long
'Mostra valores (percentuais) de comissoes na emissao / baixa

Dim lErro As Long
Dim dPercentualEm As Double
Dim dValorEmissao As Double

On Error GoTo Erro_EmissaoBaixa_Calcula
    
    'Se o Percentual de Emissao estiver Preenchido
    If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Emissao_Col))) > 0 Then
        
        'Calcula o Valor na Baixa
        dPercentualEm = CDbl(Format(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Emissao_Col), "General Number"))
        dValorEmissao = dPercentualEm * dValorComissao
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Emissao_Col) = Format(dValorEmissao, "Standard")
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Baixa_Col) = Format(dValorComissao - dValorEmissao, "Standard")
    
    'Se o Valor da Emissao estiver Preechido
    ElseIf Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Emissao_Col))) > 0 Then
            
        'Calcula valor da Baixa
        dValorEmissao = CDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Emissao_Col))

        If dValorEmissao > dValorComissao Then Error 58308

        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Valor_Baixa_Col) = Format(dValorComissao - dValorEmissao, "Standard")
        
        'Calcula Percentual da Baixa e da Emissao
        dPercentualEm = dValorEmissao / dValorComissao
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Emissao_Col) = Format(dPercentualEm, "Percent")
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Percentual_Baixa_Col) = Format(1 - dPercentualEm, "Percent")

    End If

    EmissaoBaixa_Calcula = SUCESSO

    Exit Function

Erro_EmissaoBaixa_Calcula:

    EmissaoBaixa_Calcula = Err

    Select Case Err

        Case 58308
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_COMISSAO_EMISSAO_MAIOR", Err, dValorEmissao, dValorComissao)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154410)

    End Select

    Exit Function

End Function

Private Sub OptionCupom_Click()

Dim lErro As Long

On Error GoTo Erro_OptionCupom_Click

    'Se o Documento não era CupomFiscal e foi alterado algum registro
    If iTipo_Documento <> CUPOM_FISCAL And iAlterado = REGISTRO_ALTERADO Then
        
        lErro = Teste_Salva(Me, iAlterado)
        If lErro <> SUCESSO And lErro <> 20323 Then gError 126309
            
        'Caso Cancele tem que voltar o Foco para NotaFiscal
        If lErro = 20323 Then gError 126310
        
        'Marca o Pedido e Limpa a Tela
        OptionCupom.Value = True
        Call Limpa_Tela_Comissoes
    
        iAlterado = 0
        
    'Se o Documento não era CupomFiscal
    ElseIf iTipo_Documento <> CUPOM_FISCAL Then
        
        'Limpa a tela e zera o iAlterado
        Call Limpa_Tela_Comissoes
        iAlterado = 0
    
    End If
    
    'Se o Cupom esta True Atualiza o Tipo do Documento e esconde a Série
    If OptionCupom.Value = True Then
        iTipo_Documento = CUPOM_FISCAL
        LabelECF.Visible = True
        ECF.Visible = True
        COO.Visible = True
        LabelCOO.Visible = True
        
        Codigo.Visible = False
        LabelNumero.Visible = False
        LabelSerie.Visible = False
        Serie.Visible = False
        
    End If

    Exit Sub
    
Erro_OptionCupom_Click:

    Select Case gErr

        Case 126309
        
        Case 126310
            If iTipo_Documento = NOTA_FISCAL Then NotaFiscal.Value = True
            If iTipo_Documento = PEDIDO_DE_VENDA Then PedidoVenda.Value = True
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154411)

    End Select

    Exit Sub
    
End Sub

Private Sub PedidoVenda_Click()

Dim lErro As Long

On Error GoTo Erro_PedidoVenda_Click
    
    'Se o Documento não era Pedido de Venda e foi alterado algum registro
    If iTipo_Documento <> PEDIDO_DE_VENDA And iAlterado = REGISTRO_ALTERADO Then
        
        lErro = Teste_Salva(Me, iAlterado)
        If lErro <> SUCESSO And lErro <> 20323 Then Error 58392
            
        'Caso Cancele tem que voltar o Foco para NotaFiscal
        If lErro = 20323 Then Error 58393
        
        'Marca o Pedido e Limpa a Tela
        PedidoVenda.Value = True
        Call Limpa_Tela_Comissoes
    
        iAlterado = 0
        
    'Se o Documento não era Pedido de Venda
    ElseIf iTipo_Documento <> PEDIDO_DE_VENDA Then
        
        'Limpa a tela e zera o iAlterado
        Call Limpa_Tela_Comissoes
        iAlterado = 0
    
    End If
    
    'Se o Pedidode esta True Atualiza o Tipo do Documento e esconde a Série
    If PedidoVenda.Value = True Then
        iTipo_Documento = PEDIDO_DE_VENDA
        Serie.Visible = False
        LabelSerie.Visible = False
    
        LabelECF.Visible = False
        ECF.Visible = False
        COO.Visible = False
        LabelCOO.Visible = False
        
        Codigo.Visible = True
        LabelNumero.Visible = True
    
    End If

    Exit Sub
    
Erro_PedidoVenda_Click:

    Select Case Err

        Case 58392 'Tratado na Rotina Chamada
        
        Case 58393
            If iTipo_Documento = NOTA_FISCAL Then NotaFiscal.Value = True
            If iTipo_Documento = CUPOM_FISCAL Then OptionCupom.Value = True
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154412)

    End Select

    Exit Sub

End Sub

Private Sub PedidoVenda_GotFocus()
    
    'Se o documento Já era Pedido de Venda então Recebe Pedido de Venda
    If PedidoVenda.Value = True Then
        iTipo_Documento = PEDIDO_DE_VENDA
    Else
        iTipo_Documento = NOTA_FISCAL
    End If
       
End Sub

Private Sub PercentualComissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercentualComissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub PercentualComissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub PercentualComissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = PercentualComissao
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PercentualEmissao_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercentualEmissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub PercentualEmissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub PercentualEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = PercentualEmissao
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Serie_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Serie_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Codigo_Validate

    If NotaFiscal.Value = True And Len(Trim(Serie.Text)) > 0 And Len(Trim(Codigo.Text)) > 0 Then
            
        'Verifica se ultrapassou o limite máximo da Série
        If Len(Trim(Serie.Text)) > 3 Then Error 43679
            
        objNFiscal.lNumNotaFiscal = CLng(Codigo.Text)
        objNFiscal.sSerie = LTrim(Serie.Text)
        objNFiscal.iFilialEmpresa = giFilialEmpresa
        
        lErro = CF("NF_NFFatura_Le_NumeroSerie", objNFiscal)
        If lErro <> SUCESSO And lErro <> 58324 Then Error 58330
        
        If lErro = 58324 Then Error 59737
    
    End If
    
    Exit Sub
    
Erro_Codigo_Validate:

    Cancel = True
    
    Select Case Err
                
        Case 43679
             lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_MAIOR_LIMITE_MAXIMO", Err)
        
        Case 58330 'Tratado na Rotina chamada
        
        Case 59737
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_TEM_COMISSAO", Err, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.iFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154413)

    End Select

    Exit Sub


End Sub

Private Sub ValorBaixa_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ValorBaixa_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub ValorBaixa_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub ValorBaixa_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = ValorBaixa
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ValorBase_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ValorBase_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub ValorBase_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub ValorBase_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = ValorBase
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ValorComissao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ValorComissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub ValorComissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub ValorComissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = ValorComissao
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    

End Sub

Private Sub ValorEmissao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ValorEmissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub ValorEmissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub ValorEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = ValorEmissao
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Vendedor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Vendedor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Vendedor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Vendedor
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_COMISSOES
    Set Form_Load_Ocx = Me
    Caption = "Comissões"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Comissoes"
    
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
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelNumero_Click
            
        ElseIf Me.ActiveControl Is Vendedor Then
            Call BotaoVendedores_Click
            
        ElseIf Me.ActiveControl Is ECF Then
            Call LabelECF_Click
            
        ElseIf Me.ActiveControl Is COO Then
            Call LabelCOO_Click
            
        End If
    
    End If

End Sub


Private Sub LabelSerie_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelSerie, Source, X, Y)
End Sub

Private Sub LabelSerie_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelSerie, Button, Shift, X, Y)
End Sub

Private Sub LabelNumero_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumero, Source, X, Y)
End Sub

Private Sub LabelNumero_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumero, Button, Shift, X, Y)
End Sub

Private Sub TotalPercentualComissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalPercentualComissao, Source, X, Y)
End Sub

Private Sub TotalPercentualComissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalPercentualComissao, Button, Shift, X, Y)
End Sub

Private Sub TotalValorComissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalValorComissao, Source, X, Y)
End Sub

Private Sub TotalValorComissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalValorComissao, Button, Shift, X, Y)
End Sub

'Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label11, Source, X, Y)
'End Sub

'Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
'End Sub

Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Cliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Cliente, Source, X, Y)
End Sub

Private Sub Cliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Cliente, Button, Shift, X, Y)
End Sub

Private Sub Filial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Filial, Source, X, Y)
End Sub

Private Sub Filial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Filial, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Valor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Valor, Source, X, Y)
End Sub

Private Sub Valor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Valor, Button, Shift, X, Y)
End Sub

Private Sub TotalValorBase_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalValorBase, Source, X, Y)
End Sub

Private Sub TotalValorBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalValorBase, Button, Shift, X, Y)
End Sub

