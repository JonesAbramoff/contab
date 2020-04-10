VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl FornecedorProduto 
   ClientHeight    =   5190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9120
   LockControls    =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   9120
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4050
      Index           =   2
      Left            =   210
      TabIndex        =   9
      Top             =   900
      Visible         =   0   'False
      Width           =   8670
      Begin VB.Frame Frame2 
         Caption         =   "Última Compra"
         Height          =   1665
         Left            =   150
         TabIndex        =   44
         Top             =   2265
         Width           =   8355
         Begin MSMask.MaskEdBox QuantPedida 
            Height          =   315
            Left            =   2175
            TabIndex        =   12
            Top             =   300
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantRecebida 
            Height          =   315
            Left            =   2175
            TabIndex        =   14
            Top             =   765
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Valor 
            Height          =   315
            Left            =   2175
            TabIndex        =   16
            Top             =   1215
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
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
         Begin MSComCtl2.UpDown UpDownPedido 
            Height          =   315
            Left            =   7440
            TabIndex        =   45
            Top             =   300
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataPedido 
            Height          =   315
            Left            =   6315
            TabIndex        =   13
            Top             =   300
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownReceb 
            Height          =   315
            Left            =   7440
            TabIndex        =   46
            Top             =   765
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataReceb 
            Height          =   315
            Left            =   6315
            TabIndex        =   15
            Top             =   765
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade Pedida:"
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
            Left            =   405
            TabIndex        =   51
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label12 
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
            Left            =   1590
            TabIndex        =   50
            Top             =   1230
            Width           =   510
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade Recebida:"
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
            TabIndex        =   49
            Top             =   825
            Width           =   1920
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Data do Pedido:"
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
            Left            =   4830
            TabIndex        =   48
            Top             =   360
            Width           =   1395
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Data do Recebimento:"
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
            Left            =   4305
            TabIndex        =   47
            Top             =   825
            Width           =   1920
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Pedidos de Compra"
         Height          =   1200
         Left            =   150
         TabIndex        =   40
         Top             =   945
         Width           =   8355
         Begin MSMask.MaskEdBox QuantPedAbertos 
            Height          =   315
            Left            =   4650
            TabIndex        =   10
            Top             =   300
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TempoMedio 
            Height          =   315
            Left            =   4650
            TabIndex        =   11
            Top             =   750
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tempo médio entre pedido e recebimento:"
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
            Left            =   990
            TabIndex        =   43
            Top             =   810
            Width           =   3585
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade em Pedidos Abertos:"
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
            Left            =   1785
            TabIndex        =   42
            Top             =   360
            Width           =   2790
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "dias"
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
            Left            =   5310
            TabIndex        =   41
            Top             =   810
            Width           =   360
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Produto/Fornecedor"
         Height          =   765
         Left            =   150
         TabIndex        =   33
         Top             =   75
         Width           =   8355
         Begin VB.Label UnidMed 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   3435
            TabIndex        =   39
            Top             =   300
            Width           =   900
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "U.M.:"
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
            Left            =   2910
            TabIndex        =   38
            Top             =   330
            Width           =   480
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Produto:"
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
            Left            =   210
            TabIndex        =   37
            Top             =   330
            Width           =   735
         End
         Begin VB.Label FornecedorLabel 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5820
            TabIndex        =   36
            Top             =   300
            Width           =   2295
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Fornecedor:"
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
            Left            =   4680
            TabIndex        =   35
            Top             =   330
            Width           =   1035
         End
         Begin VB.Label ProdutoLabel 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1005
            TabIndex        =   34
            Top             =   300
            Width           =   1635
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4050
      Index           =   1
      Left            =   210
      TabIndex        =   0
      Top             =   915
      Width           =   8670
      Begin VB.ListBox ListaForn 
         Height          =   3555
         IntegralHeight  =   0   'False
         Left            =   6240
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   336
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.TextBox ProdutoFornecedor 
         Height          =   315
         Left            =   2070
         MaxLength       =   20
         TabIndex        =   5
         Top             =   2352
         Width           =   1455
      End
      Begin VB.ComboBox Fornecedor 
         Height          =   315
         Left            =   2070
         TabIndex        =   3
         Top             =   1773
         Width           =   2955
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
         Height          =   195
         Left            =   5190
         TabIndex        =   4
         Top             =   1830
         Width           =   945
      End
      Begin MSComctlLib.TreeView TvwProduto 
         Height          =   3555
         Left            =   6240
         TabIndex        =   2
         Top             =   360
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   6271
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   315
         Left            =   2070
         TabIndex        =   1
         Top             =   615
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox LoteMinimo 
         Height          =   315
         Left            =   2070
         TabIndex        =   7
         Top             =   2931
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox LoteEconomico 
         Height          =   315
         Left            =   2070
         TabIndex        =   8
         Top             =   3510
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
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
         Left            =   1290
         TabIndex        =   32
         Top             =   660
         Width           =   735
      End
      Begin VB.Label LabelProduto 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   6195
         TabIndex        =   31
         Top             =   90
         Width           =   765
      End
      Begin VB.Label LabelForn 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedores"
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
         Left            =   6195
         TabIndex        =   30
         Top             =   90
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label2 
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
         Height          =   195
         Left            =   1095
         TabIndex        =   29
         Top             =   1245
         Width           =   930
      End
      Begin VB.Label Descricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2070
         TabIndex        =   28
         Top             =   1194
         Width           =   3930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor:"
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
         Left            =   990
         TabIndex        =   27
         Top             =   1815
         Width           =   1035
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Código no Fornecedor:"
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
         Left            =   75
         TabIndex        =   26
         Top             =   2400
         Width           =   1950
      End
      Begin VB.Label UnidMed 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   5100
         TabIndex        =   25
         Top             =   615
         Width           =   900
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "U.M. de Compra:"
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
         Left            =   3615
         TabIndex        =   24
         Top             =   660
         Width           =   1440
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Lote Mínimo:"
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
         Left            =   900
         TabIndex        =   23
         Top             =   2970
         Width           =   1125
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Lote Econômico:"
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
         Left            =   585
         TabIndex        =   22
         Top             =   3555
         Width           =   1440
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6780
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "FornecedorProduto.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "FornecedorProduto.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "FornecedorProduto.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "FornecedorProduto.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4455
      Left            =   150
      TabIndex        =   52
      Top             =   540
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   7858
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Complementares"
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
End
Attribute VB_Name = "FornecedorProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''responsável: Jones
''revisada em: 14/09/98
''pendencias:
'    'verificar o que deve ocorrer se ProdutoFilial nao estiver criado
'    'acho que em vez de listbox de fornecedores deveria usar browse
'
'Option Explicit
'
''Property Variables:
'Dim m_Caption As String
'Event Unload()
'
'Dim iAlterado As Integer
'Dim iProdutoAlterado As Integer
'Dim iFrameAtual As Integer
'
''Constantes públicas dos tabs
'Private Const TAB_DadosPrincipais = 1
'Private Const TAB_Complemento = 2
'
'Private Sub BotaoExcluir_Click()
'
'Dim lErro As Long
'Dim objFornecedorProduto As New ClassFornecedorProduto
'Dim objFornecedor As New ClassFornecedor
'Dim objProduto As New ClassProduto
'Dim vbMsgRes As VbMsgBoxResult
'Dim sProduto As String
'Dim iProdutoPreenchido As Integer
'
'On Error GoTo Erro_BotaoExcluir_Click
'
'    GL_objMDIForm.MousePointer = vbHourglass
'
'    'Verifica preenchimento de produto
'    If Len(Trim(Produto.Text)) = 0 Then Error 28321
'
'    'Verifica preenchimento de Fornecedor
'    If Len(Trim(Fornecedor.Text)) = 0 Then Error 28320
'
'    objFornecedor.sNomeReduzido = Fornecedor.Text
'    'Lê Fornecedor
'    lErro = CF("Fornecedor_Le_NomeReduzido",objFornecedor)
'    If lErro <> SUCESSO And lErro <> 6681 Then Error 28353
'
'    If lErro = 6681 Then Error 28354
'
'    sProduto = Produto.Text
'
'    'Critica o formato do Produto e se existe no BD
'    lErro = CF("Produto_Critica",sProduto, objProduto, iProdutoPreenchido)
'    If lErro <> SUCESSO And lErro <> 25041 Then Error 28322
'
'    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then Error 27198
'
'    If lErro = 25041 Then Error 28323
'
'   'Preenche objFornecedorProduto
'    objFornecedorProduto.lFornecedor = objFornecedor.lCodigo
'    objFornecedorProduto.sProduto = objProduto.sCodigo
'
'    lErro = CF("FornecedorProduto_Le",objFornecedorProduto)
'    If lErro <> SUCESSO And lErro <> 28240 Then Error 28324
'
'    If lErro = 28240 Then Error 28325
'
'    'Pede confirmação para exclusão Fornecedor Produto
'    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_FORNECEDOR_PRODUTO", objFornecedorProduto.lFornecedor, objFornecedorProduto.sProduto)
'
'    If vbMsgRes = vbYes Then
'
'        lErro = CF("FornecedorProduto_Exclui",objFornecedorProduto)
'        If lErro <> SUCESSO Then Error 28326
'
'        'excluir o fornecedor da combo
'        Call Combo_Item_Igual_Remove(Fornecedor)
'
'        'Limpa dados do fornecedor produto
'        Call Limpa_Campos_FornecedorProduto
'
'    End If
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Exit Sub
'
'Erro_BotaoExcluir_Click:
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Select Case Err
'
'        Case 28320
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)
'
'        Case 28321, 27198
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)
'
'        Case 28322, 28324, 28326, 28353
'
'        Case 28323
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)
'
'        Case 28325
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORPRODUTO_NAO_ENCONTRADO", Err, objFornecedorProduto.lFornecedor, objFornecedorProduto.sProduto)
'
'        Case 28354
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160637)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub BotaoFechar_Click()
'
'    Unload Me
'
'End Sub
'
'Private Sub BotaoGravar_Click()
'
'Dim lErro As Long
'Dim sForn As String
'
'On Error GoTo Erro_BotaoGravar_Click
'
'    lErro = Gravar_Registro()
'    If lErro <> SUCESSO Then Error 28278
'
'    'excluir e incluir o fornecedor na combo
'    sForn = Fornecedor.Text
'
'    FornecedorLabel = ""
'
'    Call Combo_Item_Igual_Remove(Fornecedor)
'
'    Fornecedor.Text = ""
'    Fornecedor.ListIndex = -1
'
'    Fornecedor.AddItem sForn
'
'    Call Limpa_Campos_FornecedorProduto
'
'    iAlterado = 0
'
'    Exit Sub
'
'Erro_BotaoGravar_Click:
'
'    Select Case Err
'
'        Case 28278
'            'Erro tratado na rotina chamada
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160638)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub BotaoLimpar_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_BotaoLimpar_Click
'
'    lErro = Teste_Salva(Me, iAlterado)
'    If lErro <> SUCESSO Then Error 28298
'
'    'Limpa a tela
'    Call Limpa_Tela_FornecedorProduto
'
'    Exit Sub
'
'Erro_BotaoLimpar_Click:
'
'    Select Case Err
'
'        Case 28298
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160639)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub DataPedido_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub DataPedido_GotFocus()
'
'    Call MaskEdBox_TrataGotFocus(DataPedido, iAlterado)
'
'End Sub
'
'Private Sub DataPedido_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_DataPedido_Validate
'
'    'Verifica se a data de pedido está preenchida
'    If Len(Trim(DataPedido.ClipText)) = 0 Then Exit Sub
'
'    'Verifica se a data final é válida
'    lErro = Data_Critica(DataPedido.Text)
'    If lErro <> SUCESSO Then Error 28347
'
'    Exit Sub
'
'Erro_DataPedido_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 28347
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160640)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub DataReceb_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub DataReceb_GotFocus()
'
'    Call MaskEdBox_TrataGotFocus(DataReceb, iAlterado)
'
'End Sub
'
'Private Sub DataReceb_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_DataReceb_Validate
'
'    'Verifica se a data de recebimento está preenchida
'    If Len(Trim(DataReceb.ClipText)) = 0 Then Exit Sub
'
'    'Verifica se a data final é válida
'    lErro = Data_Critica(DataReceb.Text)
'    If lErro <> SUCESSO Then Error 28348
'
'    Exit Sub
'
'Erro_DataReceb_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 28348
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160641)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub Form_Activate()
'
'    Call TelaIndice_Preenche(Me)
'
'End Sub
'
'Public Sub Form_Deactivate()
'
'    gi_ST_SetaIgnoraClick = 1
'
'End Sub
'
'Public Sub Form_Load()
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim colFornecedor As New Collection
'Dim objFornecedor As ClassFornecedor
'
'On Error GoTo Erro_Form_Load
'
'    iFrameAtual = 1
'
'    'Inicializa a mascara de produto
'    lErro = CF("Inicializa_Mascara_Produto_MaskEd",Produto)
'    If lErro <> SUCESSO Then Error 28252
'
'    'Preenche a ListBox com Fornecedores existentes no BD
'    lErro = CF("Fornecedor_Le_Todos",colFornecedor)
'    If lErro <> SUCESSO Then Error 28253
'
'    For Each objFornecedor In colFornecedor
'
'       'Insere na ListBox Código e NomeReduzido de Fornecedor
'       ListaForn.AddItem objFornecedor.lCodigo & SEPARADOR & objFornecedor.sNomeReduzido
'
'    Next
'
'    'Inicializa a Lista de Produtos
'    lErro = CF("Carga_Arvore_Produto",TvwProduto.Nodes)
'    If lErro <> SUCESSO Then Error 28254
'
'    LoteEconomico.Format = FORMATO_ESTOQUE
'    LoteMinimo.Format = FORMATO_ESTOQUE
'    QuantPedAbertos.Format = FORMATO_ESTOQUE
'    QuantPedida.Format = FORMATO_ESTOQUE
'    QuantRecebida.Format = FORMATO_ESTOQUE
'
'
'    iAlterado = 0
'
'    lErro_Chama_Tela = SUCESSO
'
'    Exit Sub
'
'Erro_Form_Load:
'
'    lErro_Chama_Tela = Err
'
'    Select Case Err
'
'        Case 28252, 28253, 28254
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160642)
'
'    End Select
'
'    iAlterado = 0
'
'    Exit Sub
'
'End Sub
'
'Function Trata_Parametros(Optional objFornecedorProduto As ClassFornecedorProduto) As Long
'
'Dim lErro As Long
'Dim sCodigo As String
'
'On Error GoTo Erro_Trata_Parametros
'
'    'Se há um FornecedorProduto selecionado, exibir seus dados
'    If Not (objFornecedorProduto Is Nothing) Then
'
'        'Verifica se o FornecedorProduto existe
'        lErro = CF("FornecedorProduto_Le",objFornecedorProduto)
'        If lErro <> SUCESSO And lErro <> 28240 Then Error 28255
'
'        'Não encontrou Fornecedor do Produto
'        If lErro = 28240 Then
'            sCodigo = objFornecedorProduto.sProduto
'            lErro = CF("Traz_Produto_MaskEd",sCodigo, Produto, Descricao)
'            If lErro <> SUCESSO Then Error 28434
'            ProdutoLabel.Caption = Produto
'            Call Produto_Validate(bSGECancelDummy)
'        Else
'            'Se encontrou Fornecedor Produto mostra os dados na tela
'            lErro = Traz_FornecedorProduto_Tela(objFornecedorProduto)
'            If lErro <> SUCESSO Then Error 28257
'        End If
'
'    End If
'
'    iAlterado = 0
'
'    Trata_Parametros = SUCESSO
'
'    Exit Function
'
'Erro_Trata_Parametros:
'
'    Trata_Parametros = Err
'
'    Select Case Err
'
'        Case 28434, 28255, 28257
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160643)
'
'    End Select
'
'    iAlterado = 0
'
'    Exit Function
'
'End Function
'
'Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
'
'    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
'
'End Sub
'
'Public Sub Form_Unload(Cancel As Integer)
'
'Dim lErro As Long
'
'  'Libera a referencia da tela e fecha o comando das setas se estiver aberto
'    lErro = ComandoSeta_Liberar(Me.Name)
'
'End Sub
'
'Private Sub Fornecedor_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub Fornecedor_Click()
'
'Dim lErro As Long, objFornecedorProduto As New ClassFornecedorProduto
'Dim objFornecedor As New ClassFornecedor, sProdutoFormatado As String, iProdutoPreenchido As Integer
'
'On Error GoTo Erro_Fornecedor_Click
'
'    iAlterado = REGISTRO_ALTERADO
'
'    If Fornecedor.ListIndex <> -1 Then
'
'        'Preencher objFornecedorProduto a partir dos campos de Produto e Fornecedor
'        'Verifica se Fornecedor existe
'        objFornecedor.sNomeReduzido = Fornecedor.Text
'        lErro = CF("Fornecedor_Le_NomeReduzido",objFornecedor)
'        If lErro <> SUCESSO And lErro <> 6681 Then Error 41516
'
'        If lErro = 6681 Then Error 41517
'
'        objFornecedorProduto.lFornecedor = objFornecedor.lCodigo
'
'        If Len(Trim(Produto.ClipText)) > 0 Then
'
'            'Critica o formato do Produto
'            lErro = CF("Produto_Formata",Produto.Text, sProdutoFormatado, iProdutoPreenchido)
'            If lErro <> SUCESSO Then Error 41518
'
'            objFornecedorProduto.sProduto = sProdutoFormatado
'
'            lErro = CF("FornecedorProduto_Le",objFornecedorProduto)
'            If lErro <> SUCESSO And lErro <> 28240 Then Error 41522
'
'            If lErro <> SUCESSO Then Error 41523
'
'            'Mostra os dados do FornecedorProduto na tela
'            lErro = Mostra_FornecedorProduto_Tela(objFornecedorProduto)
'            If lErro <> SUCESSO Then Error 40043
'
'        End If
'
'    Else
'
'        Call Limpa_Campos_FornecedorProduto
'
'    End If
'
'    Exit Sub
'
'Erro_Fornecedor_Click:
'
'    Select Case Err
'
'        Case 41523
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", Err)
'
'        Case 40043, 41516, 41518, 41522
'
'        Case 41517
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160644)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub Fornecedor_GotFocus()
'
'    LabelForn.Visible = True
'    ListaForn.Visible = True
'    TvwProduto.Visible = False
'    LabelProduto.Visible = False
'
'End Sub
'
'Private Sub Fornecedor_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objFornecedorProduto As New ClassFornecedorProduto
'Dim objFornecedor As New ClassFornecedor
'Dim objProdutoFilial As New ClassProdutoFilial
'Dim sProdutoFormatado As String
'Dim iProdutoPreenchido As Integer
'
'On Error GoTo Error_Fornecedor_Validate
'
'    'Verifica se foi preenchida a ComboBox Fornecedor
'    If Len(Trim(Fornecedor.Text)) = 0 Then Exit Sub
'
'    If Fornecedor.ListIndex = -1 Then
'
'        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
'        lErro = Combo_Item_Igual(Fornecedor)
'        If lErro <> SUCESSO And lErro <> 12253 Then Error 28264
'
'    End If
'
'    FornecedorLabel.Caption = Fornecedor.Text
'
'    Exit Sub
'
'Error_Fornecedor_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 28269
'
'        Case 28266, 28299, 28356, 28385
'
'        Case 28267
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORPRODUTO_NAO_ENCONTRADO", Err, objFornecedorProduto.lFornecedor, objFornecedorProduto.sProduto)
'            Call Limpa_Tela_FornecedorProduto
'
'        Case 28355
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160645)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Function Mostra_FornecedorProduto_Tela(objFornecedorProduto As ClassFornecedorProduto) As Long
''Mostra os dados de objFornecedorProduto, quando o produto e o fornecedor já estão selecionados
'
'Dim lErro As Long, objProdutoFilial As New ClassProdutoFilial
'
'On Error GoTo Erro_Mostra_FornecedorProduto_Tela
'
'    'Mostra os dados do Produto na tela
'    If objFornecedorProduto.sProdutoFornecedor <> "" Then
'        ProdutoFornecedor.Text = objFornecedorProduto.sProdutoFornecedor
'    Else
'        ProdutoFornecedor.Text = ""
'    End If
'
'    If objFornecedorProduto.dLoteMinimo <> 0 Then
'        LoteMinimo.Text = objFornecedorProduto.dLoteMinimo
'    Else
'        LoteMinimo.Text = ""
'    End If
'    If objFornecedorProduto.dLoteEconomico <> 0 Then
'        LoteEconomico.Text = objFornecedorProduto.dLoteEconomico
'    Else
'        LoteEconomico.Text = ""
'    End If
'    If objFornecedorProduto.dQuantPedAbertos <> 0 Then
'        QuantPedAbertos.Text = objFornecedorProduto.dQuantPedAbertos
'    Else
'        QuantPedAbertos.Text = ""
'    End If
'    If objFornecedorProduto.iTempoMedio <> 0 Then
'        TempoMedio.Text = CStr(objFornecedorProduto.iTempoMedio)
'    Else
'        TempoMedio.Text = ""
'    End If
'    If objFornecedorProduto.dQuantPedida <> 0 Then
'        QuantPedida.Text = objFornecedorProduto.dQuantPedida
'    Else
'        QuantPedida.Text = ""
'    End If
'    If objFornecedorProduto.dQuantRecebida <> 0 Then
'        QuantRecebida.Text = objFornecedorProduto.dQuantRecebida
'    Else
'        QuantRecebida.Text = ""
'    End If
'    If objFornecedorProduto.dValor <> 0 Then
'        Valor.Text = Format(objFornecedorProduto.dValor, "Fixed")
'    Else
'        Valor.Text = ""
'    End If
'    If objFornecedorProduto.dtDataPedido <> DATA_NULA Then
'        DataPedido.Text = Format(objFornecedorProduto.dtDataPedido, "dd/mm/yy")
'    Else
'        DataPedido.PromptInclude = False
'        DataPedido.Text = ""
'        DataPedido.PromptInclude = True
'    End If
'    If objFornecedorProduto.dtDataReceb <> DATA_NULA Then
'        DataReceb.Text = Format(objFornecedorProduto.dtDataReceb, "dd/mm/yy")
'    Else
'        DataReceb.PromptInclude = False
'        DataReceb.Text = ""
'        DataReceb.PromptInclude = True
'    End If
'
'    objProdutoFilial.iFilialEmpresa = giFilialEmpresa
'    objProdutoFilial.sProduto = objFornecedorProduto.sProduto
'    'Lê o ProdutoFilial
'    lErro = CF("ProdutoFilial_Le",objProdutoFilial)
'    If lErro <> SUCESSO And lErro <> 28261 Then Error 19417
'
'    If lErro <> 28261 Then
'        If objProdutoFilial.lFornecedor <> objFornecedorProduto.lFornecedor Then
'            Padrao.Value = FORN_PROD_NAO_PADRAO
'        Else
'            Padrao.Value = 1
'        End If
'    End If
'
'    iAlterado = 0
'
'    Mostra_FornecedorProduto_Tela = SUCESSO
'
'    Exit Function
'
'Erro_Mostra_FornecedorProduto_Tela:
'
'    Mostra_FornecedorProduto_Tela = Err
'
'    Select Case Err
'
'        Case 19417
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160646)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Sub ListaForn_DblClick()
'
'Dim lErro As Long
'Dim objFornecedor As New ClassFornecedor
'Dim lCodigo As Long
'Dim sListBoxItem As String
'
'On Error GoTo Erro_ListaForn_DblClick
'
'    'Se não há Fornecedor selecionado sai da rotina
'    If ListaForn.ListIndex = -1 Then Exit Sub
'
'    sListBoxItem = ListaForn.List(ListaForn.ListIndex)
'    lCodigo = LCodigo_Extrai(sListBoxItem)
'    objFornecedor.lCodigo = lCodigo
'
'    'Lê o Fornecedor
'    lErro = CF("Fornecedor_Le",objFornecedor)
'    If lErro <> SUCESSO And lErro <> 12729 Then Error 28341
'
'    'Coloca Fornecedor na ComboBox Fornecedor
'    Fornecedor.Text = objFornecedor.sNomeReduzido
'    FornecedorLabel.Caption = objFornecedor.sNomeReduzido
'    Call Fornecedor_Validate(bSGECancelDummy)
'
''    'Verifica se está preenchida com o ítem selecionado na ComboBox Fornecedor
''    If Fornecedor.Text = Fornecedor.List(Fornecedor.ListIndex) Then Exit Sub
''
''    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
''    lErro = Combo_Item_Igual(Fornecedor)
''    If lErro <> SUCESSO And lErro <> 12253 Then Error 28300
''
''    Fornecedor_Validate
''
''    'Não existe o item na ComboBox Fornecedor
''    If lErro = 12253 Then Call Limpa_Campos_FornecedorProduto
'
'    Exit Sub
'
'Erro_ListaForn_DblClick:
'
'    Select Case Err
'
'        Case 28301
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, Fornecedor.Text)
'            Fornecedor.SetFocus
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160647)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub LoteEconomico_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub LoteEconomico_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_LoteEconomico_Validate
'
'    'Verifica se algum valor foi digitado
'    If Len(Trim(LoteEconomico.ClipText)) = 0 Then Exit Sub
'
'    'Critica o valor
'    lErro = Valor_Positivo_Critica(LoteEconomico.Text)
'    If lErro <> SUCESSO Then Error 28263
'
'    Exit Sub
'
'Erro_LoteEconomico_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 28263
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160648)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub LoteMinimo_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub LoteMinimo_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_LoteMinimo_Validate
'
'    'Verifica se algum valor foi digitado
'    If Len(Trim(LoteMinimo.ClipText)) = 0 Then Exit Sub
'
'    'Critica o valor
'    lErro = Valor_Positivo_Critica(LoteMinimo.Text)
'    If lErro <> SUCESSO Then Error 28262
'
'     Exit Sub
'
'Erro_LoteMinimo_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 28262
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160649)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub Padrao_Click()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub Produto_Change()
'
'    iProdutoAlterado = 1
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub Produto_GotFocus()
'
'    LabelForn.Visible = False
'    ListaForn.Visible = False
'    TvwProduto.Visible = True
'    LabelProduto.Visible = True
'
'
'End Sub
'
'Private Sub Produto_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objProduto As New ClassProduto
'Dim colFornecedores As New Collection
'Dim vbMsgRes As VbMsgBoxResult
'Dim sProduto As String
'Dim iProdutoPreenchido As Integer
'
'On Error GoTo Erro_Produto_Validate
'
'    If iProdutoAlterado = 1 Then
'
'        'Limpa os Campos da tela
'        Fornecedor.Clear
'        FornecedorLabel.Caption = ""
'        Call Limpa_Campos_FornecedorProduto
'
'        If Len(Trim(Produto.ClipText)) > 0 Then
'
'            sProduto = Produto.Text
'
'            'Critica o formato do Produto e se existe no BD
'            lErro = CF("Produto_Critica",sProduto, objProduto, iProdutoPreenchido)
'            If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then Error 28268
'
'            If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then Error 27199
'
'            If lErro = 25041 Then Error 28269
'
'            'O Produto é gerencial
'            If lErro = 25043 Then Error 19416
'
'            'Preenche Descricao com Descrição do Produto e UnidMed
'            Descricao.Caption = objProduto.sDescricao
'            UnidMed(0).Caption = objProduto.sSiglaUMCompra
'            UnidMed(1).Caption = objProduto.sSiglaUMCompra
'
'            'carrega os Fornecedores associados a este Produto
'            lErro = Carrega_FornecedorProduto(colFornecedores, objProduto)
'            If lErro <> SUCESSO Then Error 28274
'
'            If Fornecedor.ListCount <> 0 Then Fornecedor.ListIndex = 0
'            FornecedorLabel.Caption = Fornecedor.Text
'
'        Else
'            Descricao.Caption = ""
'            UnidMed(0).Caption = ""
'            UnidMed(1).Caption = ""
'        End If
'
'        iProdutoAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_Produto_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 28268
'
'        Case 27199, 28269 'Não encontrou Produto no BD
'
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)
'
'            If vbMsgRes = vbYes Then
'                'Chama a tela de Produtos
'                Call Chama_Tela("Produtos", objProduto)
'
'            Else
'                Descricao.Caption = ""
'                UnidMed(0).Caption = ""
'                UnidMed(1).Caption = ""
'                'Segura o foco
'            End If
'
'        Case 28274
'
'        Case 19416
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160650)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub ProdutoFornecedor_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Function Carrega_FornecedorProduto(colFornecedores As Collection, objProduto As ClassProduto) As Long
''Carrega a ComboBox Fornecedor
'
'Dim lErro As Long
'Dim iIndice As Integer
'
'On Error GoTo Erro_Carrega_FornecedorProduto
'
'    'Preenche colecao de Fornecedores associados a sProduto existentes no BD
'    lErro = CF("FornecedorProduto_Le_Fornecedores",colFornecedores, objProduto)
'    If lErro <> SUCESSO Then Error 28269
'
'    For iIndice = 1 To colFornecedores.Count
'        'Preenche a ComboBox Fornecedor com os objetos da colecao colFornecedores
'        Fornecedor.AddItem colFornecedores(iIndice)
'    Next
'
'    Carrega_FornecedorProduto = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_FornecedorProduto:
'
'    Carrega_FornecedorProduto = Err
'
'    Select Case Err
'
'        Case 28269
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160651)
'
'    End Select
'
'    Exit Function
'
'End Function
'
''Private Sub ProdutoFornecedor_Validate(Cancel As Boolean)
''
''Dim lErro As Long
''
''On Error GoTo Erro_ProdutoFornecedor_Validate
''
''    'Verifica se algum valor foi digitado
''    If Len(Trim(ProdutoFornecedor.Text)) = 0 Then Exit Sub
''
''    'Critica o valor
''    lErro = Valor_Positivo_Critica(ProdutoFornecedor.Text)
''    If lErro <> SUCESSO Then Error 28340
''
''    Exit Sub
''
''Erro_ProdutoFornecedor_Validate:
'
'''    Cancel = True
'
''
''    Select Case Err
''
''        Case 28340
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160652)
''
''    End Select
''
''    Exit Sub
''
''End Sub
'
'Private Sub QuantPedAbertos_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub QuantPedAbertos_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_QuantPedAbertos_Validate
'
'    'Verifica se algum valor foi digitado
'    If Len(Trim(QuantPedAbertos.Text)) = 0 Then Exit Sub
'
'    'Critica o valor
'    lErro = Valor_Positivo_Critica(QuantPedAbertos.Text)
'    If lErro <> SUCESSO Then Error 28342
'
'    Exit Sub
'
'Erro_QuantPedAbertos_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 28342
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160653)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub QuantPedida_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub QuantPedida_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_QuantPedida_Validate
'
'    'Verifica se algum valor foi digitado
'    If Len(Trim(QuantPedida.Text)) = 0 Then Exit Sub
'
'    'Critica o valor
'    lErro = Valor_Positivo_Critica(QuantPedida.Text)
'    If lErro <> SUCESSO Then Error 28344
'
'    Exit Sub
'
'Erro_QuantPedida_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 28344
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160654)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub QuantRecebida_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub QuantRecebida_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_QuantRecebida_Validate
'
'    'Verifica se algum valor foi digitado
'    If Len(Trim(QuantRecebida.Text)) = 0 Then Exit Sub
'
'    'Critica o valor
'    lErro = Valor_Positivo_Critica(QuantRecebida.Text)
'    If lErro <> SUCESSO Then Error 28345
'
'    Exit Sub
'
'Erro_QuantRecebida_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 28345
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160655)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub TabStrip1_Click()
'
'    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
'    If TabStrip1.SelectedItem.Index <> iFrameAtual Then
'
'        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
'
'        'Torna Frame correspondente ao Tab selecionado visivel
'        Frame1(TabStrip1.SelectedItem.Index).Visible = True
'        'Torna Frame atual visivel
'        Frame1(iFrameAtual).Visible = False
'        'Armazena novo valor de iFrameAtual
'        iFrameAtual = TabStrip1.SelectedItem.Index
'
'        Select Case iFrameAtual
'
'            Case TAB_DadosPrincipais
'                Parent.HelpContextID = IDH_FORNECEDOR_PRODUTO_ID
'
'            Case TAB_Complemento
'                Parent.HelpContextID = IDH_FORNECEDOR_PRODUTO_COMPLEMENTO
'
'        End Select
'
'    End If
'
'End Sub
'
'Private Sub TempoMedio_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub TempoMedio_GotFocus()
'
'    Call MaskEdBox_TrataGotFocus(TempoMedio, iAlterado)
'
'End Sub
'
'Private Sub TempoMedio_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_TempoMedio_Validate
'
'    'Verifica se algum valor foi digitado
'    If Len(Trim(TempoMedio.Text)) = 0 Then Exit Sub
'
'    'Critica o valor
'    lErro = Valor_Positivo_Critica(TempoMedio.Text)
'    If lErro <> SUCESSO Then Error 28343
'
'    Exit Sub
'
'Erro_TempoMedio_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 28343
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160656)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub TvwProduto_Expand(ByVal objNode As MSComctlLib.Node)
'
'Dim lErro As Long
'
'On Error GoTo Erro_TvwProduto_Expand
'
'    If objNode.Tag <> NETOS_NA_ARVORE Then
'
'        'move os dados do plano de contas do banco de dados para a arvore colNodes.
'        lErro = CF("Carga_Arvore_Produto_Netos",objNode, TvwProduto.Nodes)
'        If lErro <> SUCESSO Then Error 48084
'
'    End If
'
'    Exit Sub
'
'Erro_TvwProduto_Expand:
'
'    Select Case Err
'
'        Case 48084
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160657)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub TvwProduto_NodeClick(ByVal Node As MSComctlLib.Node)
'
'Dim lErro As Long
'Dim sCodigo As String
'Dim objProduto As New ClassProduto
'Dim colFornecedores As New Collection
'
'On Error GoTo Erro_TvwProduto_NodeClick
'
'    'Armazena key do nó clicado sem caracter inicial
'    sCodigo = Right(Node.Key, Len(Node.Key) - 1)
'
'    'Verifica se produto tem filhos
'    If Node.Children > 0 Then Exit Sub
'
'    objProduto.sCodigo = sCodigo
'    'Lê Produto
'    lErro = CF("Produto_Le",objProduto)
'    If lErro <> SUCESSO And lErro <> 28030 Then Error 28275
'
'    'Verifica se Produto é gerencial
'    If objProduto.iGerencial = GERENCIAL Then Exit Sub
'
'    lErro = CF("Traz_Produto_MaskEd",sCodigo, Produto, Descricao)
'    If lErro <> SUCESSO Then Error 28276
'
'    Call Produto_Validate(bSGECancelDummy)
'
'    ProdutoLabel.Caption = Produto
'
'    'Fecha comando de setas se estiver aberto
'    lErro = ComandoSeta_Fechar(Me.Name)
'
'    iAlterado = 0
'
'    Exit Sub
'
'Erro_TvwProduto_NodeClick:
'
'    Select Case Err
'
'        Case 28275, 28276
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160658)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Function Limpa_Tela_FornecedorProduto() As Long
''Limpa a Tela FornecedorProduto
'
'Dim lErro As Long
'On Error GoTo Erro_Limpa_Tela_FornecedorProduto
'
'    'Fecha o comando das setas se estiver aberto
'    lErro = ComandoSeta_Fechar(Me.Name)
'    If lErro <> SUCESSO Then Error 40042
'
'    'Funcao generica que limpa campos da tela
'    Call Limpa_Tela(Me)
'
'    'Limpa os campos da tela que não foram limpos pela função acima
'    Descricao.Caption = ""
'    UnidMed(0).Caption = ""
'    UnidMed(1).Caption = ""
'    Fornecedor.Clear
'    FornecedorLabel.Caption = ""
'    ProdutoFornecedor.Text = ""
'    Padrao.Value = 0
'
'    iAlterado = 0
'
'    Limpa_Tela_FornecedorProduto = SUCESSO
'
'    Exit Function
'
'Erro_Limpa_Tela_FornecedorProduto:
'
'    Select Case Err
'
'        Case 40042
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160659)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Public Function Gravar_Registro() As Long
'
'Dim lErro As Long
'Dim objFornecedorProduto As New ClassFornecedorProduto
'Dim iPadrao As Integer
'
'On Error GoTo Erro_Gravar_Registro
'
'    GL_objMDIForm.MousePointer = vbHourglass
'
'    'Verifica preenchimento de Fornecedor
'    If Len(Trim(Fornecedor.Text)) = 0 Then Error 28279
'
'    'Verifica preenchimento de produto
'    If Len(Trim(Produto.ClipText)) = 0 Then Error 28280
'
'    'Chama Move_Tela_Memoria
'    lErro = Move_Tela_Memoria(objFornecedorProduto)
'    If lErro <> SUCESSO Then Error 28284
'
'    'Fornecedor padrão
'    iPadrao = Padrao.Value
'
'    'Chama FornecedorProduto_Grava
'    lErro = CF("FornecedorProduto_Grava",objFornecedorProduto, iPadrao)
'    If lErro <> SUCESSO Then Error 28285
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Gravar_Registro = SUCESSO
'
'    Exit Function
'
'Erro_Gravar_Registro:
'
'    Gravar_Registro = Err
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Select Case Err
'
'        Case 28279
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)
'
'        Case 28280
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)
'
'        Case 28284, 28285
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160660)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Move_Tela_Memoria(objFornecedorProduto As ClassFornecedorProduto) As Long
''Move os dados da Tela para objFornecedorProduto
'
'Dim lErro As Long
'Dim objProduto As New ClassProduto
'Dim objFornecedor As New ClassFornecedor
'Dim sProdutoFormatado As String
'Dim iProdutoPreenchido As Integer
'
'On Error GoTo Erro_Move_Tela_Memoria
'
'    If Len(Trim(Fornecedor.Text)) > 0 Then
'
'        'Verifica se Fornecedor existe
'        objFornecedor.sNomeReduzido = Fornecedor.Text
'        lErro = CF("Fornecedor_Le_NomeReduzido",objFornecedor)
'        If lErro <> SUCESSO And lErro <> 6681 Then Error 28286
'
'        If lErro = 6681 Then Error 28287
'
'        objFornecedorProduto.lFornecedor = objFornecedor.lCodigo
'
'    End If
'
'    If Len(Trim(Produto.ClipText)) > 0 Then
'
'        'Critica o formato do Produto
'        lErro = CF("Produto_Formata",Produto.Text, sProdutoFormatado, iProdutoPreenchido)
'        If lErro <> SUCESSO Then Error 28288
'
'        objProduto.sCodigo = sProdutoFormatado
'
'        objFornecedorProduto.sProduto = objProduto.sCodigo
'
'    End If
'
'    'Preenche objFornecedorProduto
'    objFornecedorProduto.sProdutoFornecedor = ProdutoFornecedor.Text
'
'    If Len(Trim(LoteMinimo.Text)) > 0 Then objFornecedorProduto.dLoteMinimo = CDbl(LoteMinimo.Text)
'    If Len(Trim(LoteEconomico.Text)) > 0 Then objFornecedorProduto.dLoteEconomico = CDbl(LoteEconomico.Text)
'    If Len(Trim(QuantPedAbertos.Text)) > 0 Then objFornecedorProduto.dQuantPedAbertos = CDbl(QuantPedAbertos.Text)
'    If Len(Trim(TempoMedio.Text)) > 0 Then objFornecedorProduto.iTempoMedio = CInt(TempoMedio.Text)
'    If Len(Trim(QuantPedida.Text)) > 0 Then objFornecedorProduto.dQuantPedida = CDbl(QuantPedida.Text)
'    If Len(Trim(QuantRecebida.Text)) > 0 Then objFornecedorProduto.dQuantRecebida = CDbl(QuantRecebida.Text)
'    If Len(Trim(Valor.Text)) > 0 Then objFornecedorProduto.dValor = CDbl(Valor.Text)
'
'    If Len(Trim(DataPedido.ClipText)) > 0 Then
'        objFornecedorProduto.dtDataPedido = CDate(DataPedido.Text)
'    Else
'        objFornecedorProduto.dtDataPedido = DATA_NULA
'    End If
'
'    If Len(Trim(DataReceb.ClipText)) > 0 Then
'        objFornecedorProduto.dtDataReceb = CDate(DataReceb.Text)
'    Else
'        objFornecedorProduto.dtDataReceb = DATA_NULA
'    End If
'
'    Move_Tela_Memoria = SUCESSO
'
'    Exit Function
'
'Erro_Move_Tela_Memoria:
'
'    Move_Tela_Memoria = Err
'
'    Select Case Err
'
'        Case 28286, 28288
'
'        Case 28287
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160661)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Traz_FornecedorProduto_Tela(objFornecedorProduto As ClassFornecedorProduto) As Long
''Traz os dados da FornecedorProduto para a Tela
'
'Dim lErro As Long
'Dim sCodigo As String
'Dim objFornecedor As New ClassFornecedor
'
'On Error GoTo Erro_Traz_FornecedorProduto_Tela
'
'    sCodigo = objFornecedorProduto.sProduto
'
'    lErro = CF("Traz_Produto_MaskEd",sCodigo, Produto, Descricao)
'    If lErro <> SUCESSO Then Error 28291
'
'    objFornecedor.lCodigo = objFornecedorProduto.lFornecedor
'    'Lê o Fornecedor
'    lErro = CF("Fornecedor_Le",objFornecedor)
'    If lErro <> SUCESSO And lErro <> 12729 Then Error 28292
'
'    'Mostra Fornecedor na Tela
'    Fornecedor.Text = objFornecedor.sNomeReduzido
'    FornecedorLabel.Caption = objFornecedor.sNomeReduzido
'
'    Call Combo_Item_Igual(Fornecedor)
'
''    'Lê FornecedorProduto
''    lErro = CF("FornecedorProduto_Le",objFornecedorProduto)
''    If lErro <> SUCESSO And lErro <> 28240 Then Error 28293
''
''    'Mostra os dados de FornecedorProduto na Tela
''    lErro = Mostra_FornecedorProduto_Tela(objFornecedorProduto, objProdutoFilial)
''    If lErro <> SUCESSO Then Error yyyy
'
'    iAlterado = 0
'
'    Traz_FornecedorProduto_Tela = SUCESSO
'
'    Exit Function
'
'Erro_Traz_FornecedorProduto_Tela:
'
'    Traz_FornecedorProduto_Tela = Err
'
'    Select Case Err
'
'        Case 28291, 28292, 28294, 28295
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160662)
'
'    End Select
'
'    Exit Function
'
'End Function
'
''Function Traz_Dados_FornecedorProduto(objFornecedorProduto As ClassFornecedorProduto) As Long
'''Mostra os dados de FornecedorProduto na Tela
''Dim lErro As Long
''Dim objFornecedor As New ClassFornecedor
''Dim objProdutoFilial As New ClassProdutoFilial
''Dim objProduto As New ClassProduto
''
''On Error GoTo Erro_Traz_Dados_FornecedorProduto
''
''    'Lê FornecedorProduto
''    lErro = CF("FornecedorProduto_Le",objFornecedorProduto)
''    If lErro <> SUCESSO And lErro <> 28240 Then Error 41513
''
''    ProdutoFornecedor.Text = objFornecedorProduto.sProdutoFornecedor
''
''    If objFornecedorProduto.dLoteMinimo <> 0 Then
''        LoteMinimo.Text = CStr(objFornecedorProduto.dLoteMinimo)
''    Else
''        LoteMinimo.Text = ""
''    End If
''
''    If objFornecedorProduto.dLoteEconomico <> 0 Then
''        LoteEconomico.Text = CStr(objFornecedorProduto.dLoteEconomico)
''    Else
''        LoteEconomico.Text = ""
''    End If
''
''    If objFornecedorProduto.dQuantPedAbertos <> 0 Then
''        QuantPedAbertos.Text = CStr(objFornecedorProduto.dQuantPedAbertos)
''    Else
''        QuantPedAbertos.Text = ""
''    End If
''
''    If objFornecedorProduto.iTempoMedio <> 0 Then
''        TempoMedio.Text = CStr(objFornecedorProduto.iTempoMedio)
''    Else
''        TempoMedio.Text = ""
''    End If
''
''    If objFornecedorProduto.dQuantPedida <> 0 Then
''        QuantPedida.Text = CStr(objFornecedorProduto.dQuantPedida)
''    Else
''        QuantPedida.Text = ""
''    End If
''
''    If objFornecedorProduto.dQuantRecebida <> 0 Then
''        QuantRecebida.Text = CStr(objFornecedorProduto.dQuantRecebida)
''    Else
''        QuantRecebida.Text = ""
''    End If
''
''    If objFornecedorProduto.dValor <> 0 Then
''        Valor.Text = Format(objFornecedorProduto.dQuantPedida, "Fixed")
''    Else
''        Valor.Text = ""
''    End If
''
''    If objFornecedorProduto.dtDataPedido <> DATA_NULA Then
''        DataPedido.Text = Format(objFornecedorProduto.dtDataPedido, "dd/mm/yy")
''    Else
''        DataPedido.PromptInclude = False
''        DataPedido.Text = ""
''        DataPedido.PromptInclude = True
''    End If
''
''    If objFornecedorProduto.dtDataReceb <> DATA_NULA Then
''        DataReceb.Text = Format(objFornecedorProduto.dtDataReceb, "dd/mm/yy")
''    Else
''        DataReceb.PromptInclude = False
''        DataReceb.Text = ""
''        DataReceb.PromptInclude = True
''    End If
''
''    objProdutoFilial.sProduto = objFornecedorProduto.sProduto
''    objProdutoFilial.iFilialEmpresa = giFilialEmpresa
''    'Lê o Fornecedor padrão
''    lErro = CF("ProdutoFilial_Le",objProdutoFilial)
''    If lErro <> SUCESSO And lErro <> 28261 Then Error 41514
''
''    If lErro <> 28261 Then
''        If objProdutoFilial.lFornecedor <> objFornecedorProduto.lFornecedor Then
''            Padrao.Value = FORN_PROD_NAO_PADRAO
''        Else
''            Padrao.Value = 1
''        End If
''    End If
''
''    objProduto.sCodigo = objFornecedorProduto.sProduto
''
''    'Lê Produto
''    lErro = CF("Produto_Le",objProduto)
''    If lErro <> SUCESSO And lErro <> 28030 Then Error 28295
''
''    UnidMed(0).Caption = objProduto.sSiglaUMCompra
''    UnidMed(1).Caption = objProduto.sSiglaUMCompra
''
''    Traz_Dados_FornecedorProduto = SUCESSO
''
''    Exit Function
''
''Erro_Traz_Dados_FornecedorProduto:
''
''    Traz_Dados_FornecedorProduto = Err
''
''    Select Case Err
''
''        Case 41513, 41514, 28295
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160663)
''
''    End Select
''
''    Exit Function
''
''End Function
'
'
''""""""""""""""""""""""""""""""""""""""""""""""
''"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
''""""""""""""""""""""""""""""""""""""""""""""""
'
'Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
''Extrai os campos da tela que correspondem aos campos no BD
'
'Dim lErro As Long
'Dim objFornecedorProduto As New ClassFornecedorProduto
'Dim iIndice As Integer
'
'On Error GoTo Erro_Tela_Extrai
'
'    'Informa tabela associada à Tela
'    sTabela = "FornecedorProduto"
'
'    'Lê os dados da Tela FornecedorProduto
'    lErro = Move_Tela_Memoria(objFornecedorProduto)
'    If lErro <> SUCESSO Then Error 28296
'
'    'Preenche a coleção colCampoValor, com nome do campo,
'    'valor atual (com a tipagem do BD), tamanho do campo
'    'no BD no caso de STRING e Key igual ao nome do campo
'    colCampoValor.Add "Fornecedor", objFornecedorProduto.lFornecedor, 0, "Fornecedor"
'    colCampoValor.Add "Produto", objFornecedorProduto.sProduto, STRING_PRODUTO, "Produto"
'    colCampoValor.Add "ProdutoFornecedor", objFornecedorProduto.sProdutoFornecedor, STRING_PRODUTO_FORNECEDOR, "ProdutoFornecedor"
'    colCampoValor.Add "LoteMinimo", objFornecedorProduto.dLoteMinimo, 0, "LoteMinimo"
'    colCampoValor.Add "LoteEconomico", objFornecedorProduto.dLoteEconomico, 0, "LoteEconomico"
'    colCampoValor.Add "QuantPedAbertos", objFornecedorProduto.dQuantPedAbertos, 0, "QuantPedAbertos"
'    colCampoValor.Add "TempoMedio", objFornecedorProduto.iTempoMedio, 0, "TempoMedio"
'    colCampoValor.Add "QuantPedida", objFornecedorProduto.dQuantPedida, 0, "QuantPedida"
'    colCampoValor.Add "QuantRecebida", objFornecedorProduto.dQuantRecebida, 0, "QuantRecebida"
'    colCampoValor.Add "Valor", objFornecedorProduto.dValor, 0, "Valor"
'    colCampoValor.Add "DataPedido", objFornecedorProduto.dtDataPedido, 0, "DataPedido"
'    colCampoValor.Add "DataReceb", objFornecedorProduto.dtDataReceb, 0, "DataReceb"
'
'    Exit Sub
'
'Erro_Tela_Extrai:
'
'    Select Case Err
'
'        Case 28296
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160664)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
''Preenche os campos da tela com os correspondentes do BD
'
'Dim lErro As Long
'Dim objFornecedorProduto As New ClassFornecedorProduto
'
'On Error GoTo Erro_Tela_Preenche
'
'    objFornecedorProduto.lFornecedor = colCampoValor.Item("Fornecedor").vValor
'    objFornecedorProduto.sProduto = colCampoValor.Item("Produto").vValor
'
'    If (objFornecedorProduto.lFornecedor <> 0) And (objFornecedorProduto.sProduto <> "") Then
'
''        'Carrega objFornecedorProduto com os dados passados em colCampoValor
''        objFornecedorProduto.sProdutoFornecedor = colCampoValor.Item("ProdutoFornecedor").vValor
''        objFornecedorProduto.dLoteMinimo = colCampoValor.Item("LoteMinimo").vValor
''        objFornecedorProduto.dLoteEconomico = colCampoValor.Item("LoteEconomico").vValor
''        objFornecedorProduto.dQuantPedAbertos = colCampoValor.Item("QuantPedAbertos").vValor
''        objFornecedorProduto.iTempoMedio = colCampoValor.Item("TempoMedio").vValor
''        objFornecedorProduto.dQuantPedida = colCampoValor.Item("QuantPedida").vValor
''        objFornecedorProduto.dQuantRecebida = colCampoValor.Item("QuantRecebida").vValor
''        objFornecedorProduto.dValor = colCampoValor.Item("Valor").vValor
''        objFornecedorProduto.dtDataPedido = colCampoValor.Item("DataPedido").vValor
''        objFornecedorProduto.dtDataReceb = colCampoValor.Item("DataReceb").vValor
'
'        'Traz os dados para tela
'        lErro = Traz_FornecedorProduto_Tela(objFornecedorProduto)
'        If lErro <> SUCESSO Then Error 28297
'
'    End If
'
'    Exit Sub
'
'Erro_Tela_Preenche:
'
'    Select Case Err
'
'        Case 28297
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160665)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub UpDownPedido_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub UpDownPedido_DownClick()
'
'Dim lErro As Long
'Dim sData As String
'
'On Error GoTo Erro_UpDownPedido_DownClick
'
'    DataPedido.SetFocus
'
'    If Len(DataPedido.ClipText) > 0 Then
'
'        sData = DataPedido.Text
'
'        lErro = Data_Diminui(sData)
'        If lErro <> SUCESSO Then Error 28349
'
'        DataPedido.Text = sData
'
'    End If
'
'    Exit Sub
'
'Erro_UpDownPedido_DownClick:
'
'    Select Case Err
'
'        Case 28349
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160666)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub UpDownPedido_UpClick()
'
'Dim lErro As Long
'Dim sData As String
'
'On Error GoTo Erro_UpDownPedido_UpClick
'
'    DataPedido.SetFocus
'
'    If Len(DataPedido.ClipText) > 0 Then
'
'        sData = DataPedido.Text
'
'        lErro = Data_Aumenta(sData)
'        If lErro <> SUCESSO Then Error 28350
'
'        DataPedido.Text = sData
'
'    End If
'
'    Exit Sub
'
'Erro_UpDownPedido_UpClick:
'
'    Select Case Err
'
'        Case 28350
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160667)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub UpDownReceb_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub UpDownReceb_DownClick()
'
'Dim lErro As Long
'Dim sData As String
'
'On Error GoTo Erro_UpDownReceb_DownClick
'
'    DataPedido.SetFocus
'
'    If Len(DataReceb.ClipText) > 0 Then
'
'        sData = DataReceb.Text
'
'        lErro = Data_Diminui(sData)
'        If lErro <> SUCESSO Then Error 28351
'
'        DataReceb.Text = sData
'
'    End If
'
'    Exit Sub
'
'Erro_UpDownReceb_DownClick:
'
'    Select Case Err
'
'        Case 28351
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160668)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub UpDownReceb_UpClick()
'
'Dim lErro As Long
'Dim sData As String
'
'On Error GoTo Erro_UpDownReceb_UpClick
'
'    DataReceb.SetFocus
'
'    If Len(DataReceb.ClipText) > 0 Then
'
'        sData = DataReceb.Text
'
'        lErro = Data_Aumenta(sData)
'        If lErro <> SUCESSO Then Error 28352
'
'        DataReceb.Text = sData
'
'    End If
'
'    Exit Sub
'
'Erro_UpDownReceb_UpClick:
'
'    Select Case Err
'
'        Case 28352
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160669)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub Valor_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub Valor_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_Valor_Validate
'
'    'Verifica se algum valor foi digitado
'    If Len(Trim(Valor.ClipText)) = 0 Then Exit Sub
'
'    'Critica o valor
'    lErro = Valor_NaoNegativo_Critica(Valor.Text)
'    If lErro <> SUCESSO Then Error 28346
'
'    'Põe o valor formatado na tela
'    Valor.Text = Format(Valor.Text, "Fixed")
'
'    Exit Sub
'
'Erro_Valor_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 28346
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160670)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub Limpa_Campos_FornecedorProduto()
'
'    Padrao.Value = 0
'
'    'Limpa campos da tela
'    ProdutoFornecedor.Text = ""
'    LoteMinimo.Text = ""
'    LoteEconomico.Text = ""
'    QuantPedAbertos.Text = ""
'    TempoMedio.Text = ""
'    QuantPedida.Text = ""
'    QuantRecebida.Text = ""
'    Valor.Text = ""
'
'    DataPedido.PromptInclude = False
'    DataPedido.Text = ""
'    DataPedido.PromptInclude = True
'
'    DataReceb.PromptInclude = False
'    DataReceb.Text = ""
'    DataReceb.PromptInclude = True
'
'End Sub
'
''**** inicio do trecho a ser copiado *****
'
'Public Function Form_Load_Ocx() As Object
'
'    Parent.HelpContextID = IDH_FORNECEDOR_PRODUTO_ID
'    Set Form_Load_Ocx = Me
'    Caption = "Produto X Fornecedor"
'    Call Form_Load
'
'End Function
'
'Public Function Name() As String
'
'    Name = "FornecedorProduto"
'
'End Function
'
'Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
'End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,Controls
'Public Property Get Controls() As Object
'    Set Controls = UserControl.Controls
'End Property
'
'Public Property Get hWnd() As Long
'    hWnd = UserControl.hWnd
'End Property
'
'Public Property Get Height() As Long
'    Height = UserControl.Height
'End Property
'
'Public Property Get Width() As Long
'    Width = UserControl.Width
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,ActiveControl
'Public Property Get ActiveControl() As Object
'    Set ActiveControl = UserControl.ActiveControl
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,Enabled
'Public Property Get Enabled() As Boolean
'    Enabled = UserControl.Enabled
'End Property
'
'Public Property Let Enabled(ByVal New_Enabled As Boolean)
'    UserControl.Enabled() = New_Enabled
'    PropertyChanged "Enabled"
'End Property
'
''Load property values from storage
'Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'
'    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
'End Sub
'
''Write property values to storage
'Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'
'    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
'End Sub
'
'Private Sub Unload(objme As Object)
'
'   RaiseEvent Unload
'
'End Sub
'
'Public Property Get Caption() As String
'    Caption = m_Caption
'End Property
'
'Public Property Let Caption(ByVal New_Caption As String)
'    Parent.Caption = New_Caption
'    m_Caption = New_Caption
'End Property
'
''**** fim do trecho a ser copiado *****
'
'
'Private Sub UnidMed_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(UnidMed(Index), Source, X, Y)
'End Sub
'
'Private Sub UnidMed_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(UnidMed(Index), Button, Shift, X, Y)
'End Sub
'
'
'Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label1, Source, X, Y)
'End Sub
'
'Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
'End Sub
'
'Private Sub LabelProduto_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelProduto, Source, X, Y)
'End Sub
'
'Private Sub LabelProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelProduto, Button, Shift, X, Y)
'End Sub
'
'Private Sub LabelForn_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelForn, Source, X, Y)
'End Sub
'
'Private Sub LabelForn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelForn, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label2, Source, X, Y)
'End Sub
'
'Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
'End Sub
'
'Private Sub Descricao_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Descricao, Source, X, Y)
'End Sub
'
'Private Sub Descricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Descricao, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label3, Source, X, Y)
'End Sub
'
'Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label8, Source, X, Y)
'End Sub
'
'Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label11, Source, X, Y)
'End Sub
'
'Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label5, Source, X, Y)
'End Sub
'
'Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label10, Source, X, Y)
'End Sub
'
'Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label7, Source, X, Y)
'End Sub
'
'Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label12, Source, X, Y)
'End Sub
'
'Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label13, Source, X, Y)
'End Sub
'
'Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label9, Source, X, Y)
'End Sub
'
'Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label14, Source, X, Y)
'End Sub
'
'Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label6, Source, X, Y)
'End Sub
'
'Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label4, Source, X, Y)
'End Sub
'
'Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label15, Source, X, Y)
'End Sub
'
'Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label17, Source, X, Y)
'End Sub
'
'Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label18, Source, X, Y)
'End Sub
'
'Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
'End Sub
'
'Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
'End Sub
'
'Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label16, Source, X, Y)
'End Sub
'
'Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
'End Sub
'
'Private Sub ProdutoLabel_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(ProdutoLabel, Source, X, Y)
'End Sub
'
'Private Sub ProdutoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(ProdutoLabel, Button, Shift, X, Y)
'End Sub
'
'
'Private Sub TabStrip1_BeforeClick(Cancel As Integer)
'    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
'End Sub
'
