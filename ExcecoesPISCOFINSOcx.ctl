VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ExcecoesPISCOFINSOcx 
   ClientHeight    =   6690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   9510
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7305
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   30
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ExcecoesPISCOFINSOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ExcecoesPISCOFINSOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ExcecoesPISCOFINSOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ExcecoesPISCOFINSOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Critério"
      Height          =   2190
      Left            =   45
      TabIndex        =   26
      Top             =   540
      Width           =   9405
      Begin VB.Frame Frame3 
         Caption         =   "Fornecedores"
         Height          =   885
         Index           =   1
         Left            =   90
         TabIndex        =   46
         Top             =   1230
         Visible         =   0   'False
         Width           =   9210
         Begin VB.ComboBox CategoriaFornecedor 
            Height          =   315
            Left            =   2820
            TabIndex        =   49
            Top             =   165
            Width           =   6300
         End
         Begin VB.ComboBox ItemCategoriaFornecedor 
            Height          =   315
            Left            =   3180
            TabIndex        =   48
            Top             =   510
            Width           =   5940
         End
         Begin VB.CheckBox TodosFornecedores 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   720
            TabIndex        =   47
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label7 
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
            Left            =   2625
            TabIndex        =   51
            Top             =   525
            Width           =   510
         End
         Begin VB.Label Label5 
            Caption         =   "Categoria:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1845
            TabIndex        =   50
            Top             =   210
            Width           =   930
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Clientes"
         Height          =   870
         Index           =   0
         Left            =   90
         TabIndex        =   27
         Top             =   1230
         Width           =   9210
         Begin VB.CheckBox TodosClientes 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   720
            TabIndex        =   4
            Top             =   180
            Width           =   915
         End
         Begin VB.ComboBox ItemCategoriaCliente 
            Height          =   315
            Left            =   3180
            TabIndex        =   6
            Top             =   495
            Width           =   5940
         End
         Begin VB.ComboBox CategoriaCliente 
            Height          =   315
            Left            =   2805
            TabIndex        =   5
            Top             =   150
            Width           =   6300
         End
         Begin VB.Label Label4 
            Caption         =   "Categoria:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1845
            TabIndex        =   29
            Top             =   210
            Width           =   930
         End
         Begin VB.Label Label6 
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
            Left            =   2625
            TabIndex        =   28
            Top             =   510
            Width           =   510
         End
      End
      Begin VB.OptionButton OptCliForn 
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
         Height          =   240
         Index           =   0
         Left            =   1635
         TabIndex        =   45
         Top             =   1020
         Value           =   -1  'True
         Width           =   1920
      End
      Begin VB.OptionButton OptCliForn 
         Caption         =   "Fornecedor"
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
         Index           =   1
         Left            =   3735
         TabIndex        =   44
         Top             =   1020
         Width           =   1920
      End
      Begin VB.Frame Frame4 
         Caption         =   "Produtos"
         Height          =   840
         Left            =   105
         TabIndex        =   30
         Top             =   180
         Width           =   9210
         Begin VB.ComboBox CategoriaProduto 
            Height          =   315
            Left            =   2835
            TabIndex        =   2
            Top             =   135
            Width           =   6300
         End
         Begin VB.ComboBox ItemCategoriaProduto 
            Height          =   315
            Left            =   3195
            TabIndex        =   3
            Top             =   465
            Width           =   5940
         End
         Begin VB.CheckBox TodosProdutos 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   720
            TabIndex        =   1
            Top             =   165
            Width           =   915
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2610
            TabIndex        =   32
            Top             =   495
            Width           =   510
         End
         Begin VB.Label Label1 
            Caption         =   "Categoria:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   1860
            TabIndex        =   31
            Top             =   195
            Width           =   930
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tratamento"
      Height          =   3915
      Index           =   1
      Left            =   45
      TabIndex        =   25
      Top             =   2715
      Width           =   9405
      Begin VB.Frame FrameCOFINS 
         Caption         =   "COFINS"
         Height          =   1575
         Left            =   120
         TabIndex        =   39
         Top             =   2235
         Width           =   9210
         Begin VB.ComboBox COFINSTipoTributacaoE 
            Height          =   315
            Left            =   1665
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   525
            Width           =   7425
         End
         Begin VB.ComboBox COFINSTipoTributacao 
            Height          =   315
            Left            =   1665
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   180
            Width           =   7425
         End
         Begin VB.ComboBox COFINSTipoCalculo 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1665
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   870
            Width           =   4005
         End
         Begin MSMask.MaskEdBox COFINSAliquotaRS 
            Height          =   285
            Left            =   4590
            TabIndex        =   19
            Top             =   1215
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox COFINSAliquota 
            Height          =   285
            Left            =   1665
            TabIndex        =   18
            Top             =   1215
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "##0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            Caption         =   "Variação Entrada:"
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
            Height          =   285
            Index           =   9
            Left            =   75
            TabIndex        =   54
            Top             =   570
            Width           =   1635
         End
         Begin VB.Label Label1 
            Caption         =   "Classificação:"
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
            Height          =   285
            Index           =   6
            Left            =   435
            TabIndex        =   43
            Top             =   225
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Alíquota (R$):"
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
            Left            =   3345
            TabIndex        =   42
            Top             =   1260
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Alíquota (%):"
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
            Index           =   4
            Left            =   540
            TabIndex        =   41
            Top             =   1245
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cálculo:"
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
            TabIndex        =   40
            Top             =   930
            Width           =   1395
         End
      End
      Begin VB.Frame FramePIS 
         Caption         =   "PIS"
         Height          =   1560
         Left            =   120
         TabIndex        =   35
         Top             =   660
         Width           =   9210
         Begin VB.ComboBox PISTipoTributacao 
            Height          =   315
            Left            =   1665
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   180
            Width           =   7425
         End
         Begin VB.ComboBox PISTipoTributacaoE 
            Height          =   315
            Left            =   1665
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   525
            Width           =   7425
         End
         Begin VB.ComboBox PISTipoCalculo 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1665
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   870
            Width           =   4005
         End
         Begin MSMask.MaskEdBox PISAliquotaRS 
            Height          =   285
            Left            =   4590
            TabIndex        =   14
            Top             =   1215
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PISAliquota 
            Height          =   285
            Left            =   1665
            TabIndex        =   13
            Top             =   1215
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "##0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            Caption         =   "Classificação:"
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
            Height          =   285
            Index           =   3
            Left            =   450
            TabIndex        =   53
            Top             =   225
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "Variação Entrada:"
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
            Height          =   285
            Index           =   8
            Left            =   75
            TabIndex        =   52
            Top             =   570
            Width           =   1635
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cálculo:"
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
            Index           =   7
            Left            =   240
            TabIndex        =   38
            Top             =   930
            Width           =   1395
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Alíquota (%):"
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
            Index           =   1
            Left            =   540
            TabIndex        =   37
            Top             =   1245
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Alíquota (R$):"
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
            Index           =   2
            Left            =   3345
            TabIndex        =   36
            Top             =   1260
            Width           =   1200
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Válido para"
         Height          =   405
         Left            =   120
         TabIndex        =   34
         Top             =   225
         Width           =   9210
         Begin VB.OptionButton optPISCOFINS 
            Caption         =   "Somente COFINS"
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
            Left            =   5730
            TabIndex        =   9
            Top             =   165
            Width           =   1965
         End
         Begin VB.OptionButton optPISCOFINS 
            Caption         =   "Somente PIS"
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
            Left            =   3645
            TabIndex        =   8
            Top             =   165
            Width           =   1560
         End
         Begin VB.OptionButton optPISCOFINS 
            Caption         =   "PIS e COFINS"
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
            Left            =   1485
            TabIndex        =   7
            Top             =   150
            Value           =   -1  'True
            Width           =   1560
         End
      End
   End
   Begin VB.TextBox Fundamentacao 
      Height          =   288
      Left            =   1650
      TabIndex        =   0
      Top             =   240
      Width           =   5100
   End
   Begin VB.Label LabelFundamentacao 
      Caption         =   "Fundamentação:"
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
      Height          =   240
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   33
      Top             =   270
      Width           =   1440
   End
End
Attribute VB_Name = "ExcecoesPISCOFINSOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim colTiposTribPisCofins As New Collection

Private WithEvents objEventoExcPisCofins As AdmEvento
Attribute objEventoExcPisCofins.VB_VarHelpID = -1

Private Sub Traz_Excecao_Tela(objExcPisCofins As ClassPISCOFINSExcecao)
'Preenche a Tela

Dim lErro As Long
Dim iIndice As Integer, iCodigo As Integer
Dim objTipoTribPISCOFINS As New ClassTipoTribPISCOFINS
Dim bCancel As Boolean

On Error GoTo Erro_Traz_Excecao_Tela

    lErro = CF("PISCOFINSExcecoes_Le", objExcPisCofins)
    If lErro <> SUCESSO Then gError 205434
       
    'Preenche a Fundamentação
    Fundamentacao.Text = objExcPisCofins.sFundamentacao
    
    TodosProdutos.Value = vbUnchecked
    TodosClientes.Value = vbUnchecked
    
    'Se a Categoria do Produto estiver Preenchida
    If objExcPisCofins.sCategoriaProduto <> "" Then
        
        'Coloca na Tela e chaama o validate
        CategoriaProduto.Text = objExcPisCofins.sCategoriaProduto
        Call CategoriaProduto_Validate(bCancel)
        
        'Coloca o ItemCategoriaProduto na tela e chama o lostFocus
        ItemCategoriaProduto.Text = objExcPisCofins.sCategoriaProdutoItem
        Call ItemCategoriaProduto_Validate(bSGECancelDummy)

    Else
        'Senão marca a Check Todos
        TodosProdutos.Value = 1
    End If

    OptCliForn(objExcPisCofins.iTipoCliForn).Value = True
    Call OptCliForn_Click(objExcPisCofins.iTipoCliForn)
    
    If objExcPisCofins.iTipoCliForn = PISCOFINSEXCECOES_TIPOCLIFORN_CLIENTE Then
    
        'Se a Categoria do Cliente estiver Preenchida
        If objExcPisCofins.sCategoriaCliente <> "" Then
            
            'Coloca na Tela e chaama o validate
            CategoriaCliente.Text = objExcPisCofins.sCategoriaCliente
            Call CategoriaCliente_Validate(bCancel)
    
            'Coloca o ItemCategoriaCliente na tela e chama o lostFocus
            ItemCategoriaCliente.Text = objExcPisCofins.sCategoriaClienteItem
            Call ItemCategoriaCliente_Validate(bSGECancelDummy)
    
        Else
            'Senão marca a Check Todos
            TodosClientes.Value = 1
        End If
    
    Else
        
        'Se a Categoria do Fornecedor estiver Preenchida
        If objExcPisCofins.sCategoriaFornecedor <> "" Then
            
            'Coloca na Tela e chaama o validate
            CategoriaFornecedor.Text = objExcPisCofins.sCategoriaFornecedor
            Call CategoriaFornecedor_Validate(bCancel)
    
            'Coloca o ItemCategoriaFornecedor na tela e chama o lostFocus
            ItemCategoriaFornecedor.Text = objExcPisCofins.sCategoriaFornecedorItem
            Call ItemCategoriaFornecedor_Validate(bSGECancelDummy)
    
        Else
            'Senão marca a Check Todos
            TodosFornecedores.Value = 1
        End If
    
    End If
    
    optPISCOFINS.Item(objExcPisCofins.iTipo).Value = True
    Call optPISCOFINS_Click(objExcPisCofins.iTipo)

    Call Combo_Seleciona_ItemData(PISTipoTributacao, objExcPisCofins.iTipoPIS)
    Call PISTipoTributacao_Click
    
    If objExcPisCofins.iTipoPISE <> 0 Then Call Combo_Seleciona_ItemData(PISTipoTributacaoE, objExcPisCofins.iTipoPISE)
        
    Call Combo_Seleciona_ItemData(PISTipoCalculo, objExcPisCofins.iPISTipoCalculo)
    Call PISTipoCalculo_Click
    
    If PISAliquota.Enabled = True Then PISAliquota.Text = CStr(objExcPisCofins.dAliquotaPisPerc * 100)
    If PISAliquotaRS.Enabled = True Then PISAliquotaRS.Text = Format(objExcPisCofins.dAliquotaPisRS, PISAliquotaRS.Format)

    Call Combo_Seleciona_ItemData(COFINSTipoTributacao, objExcPisCofins.iTipoCOFINS)
    Call COFINSTipoTributacao_Click
    
    If objExcPisCofins.iTipoCOFINSE <> 0 Then Call Combo_Seleciona_ItemData(COFINSTipoTributacaoE, objExcPisCofins.iTipoCOFINSE)
    
    Call Combo_Seleciona_ItemData(COFINSTipoCalculo, objExcPisCofins.iCOFINSTipoCalculo)
    Call COFINSTipoCalculo_Click
    
    If COFINSAliquota.Enabled = True Then COFINSAliquota.Text = CStr(objExcPisCofins.dAliquotaCofinsPerc * 100)
    If COFINSAliquotaRS.Enabled = True Then COFINSAliquotaRS.Text = Format(objExcPisCofins.dAliquotaCofinsRS, COFINSAliquotaRS.Format)

    Call optPISCOFINS_Click(objExcPisCofins.iTipo)
    
    Call COFINSTipoTributacao_Click
    Call PISTipoTributacao_Click

    iAlterado = 0

    Exit Sub

Erro_Traz_Excecao_Tela:

    Select Case gErr
    
        Case 205434

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205419)

    End Select

    Exit Sub

End Sub

Private Sub Aliquota_Validate(objControle As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Aliquota_Validate

    If Len(objControle.Text) > 0 Then

        'Testa o valor
        lErro = Porcentagem_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 205420

    End If

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_Aliquota_Validate:

    Cancel = True

    Select Case gErr

        Case 205420

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205421)

    End Select

    Exit Sub

End Sub

Public Sub Valor_Validate(objControle As Object, Cancel As Boolean)

Dim lErro As Long
Dim dValor As Double

On Error GoTo Erro_Valor_Validate

    If Len(Trim(objControle.Text)) > 0 Then

        'Critica se valor é não negativo
        lErro = Valor_NaoNegativo_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 198580

        dValor = CDbl(objControle.Text)

        objControle.Text = Format(dValor, objControle.Format)

    End If

    Exit Sub

Erro_Valor_Validate:

    Cancel = True

    Select Case gErr

        Case 205420

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205421)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaCliente_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objCategoriaClienteItem As ClassCategoriaClienteItem
Dim objCategoriaCliente As New ClassCategoriaCliente
Dim colCategoria As New Collection

On Error GoTo Erro_CategoriaCliente_Click

    iAlterado = REGISTRO_ALTERADO

    'Verifica se a CategoriaCliente foi preenchida
    If CategoriaCliente.ListIndex <> -1 Then

        objCategoriaCliente.sCategoria = CategoriaCliente.Text

        'Lê os dados de Itens da Categoria do Cliente
        lErro = CF("CategoriaCliente_Le_Itens", objCategoriaCliente, colCategoria)
        If lErro <> SUCESSO Then gError 205422

        ItemCategoriaCliente.Enabled = True

        'Limpa os dados de ItemCategoriaCliente
        ItemCategoriaCliente.Clear

        'Preenche ItemCategoriaCliente
        For Each objCategoriaClienteItem In colCategoria

            ItemCategoriaCliente.AddItem objCategoriaClienteItem.sItem

        Next
        TodosClientes.Value = 0

    Else
        
        'Senão Desabilita e limpa ItemCategoriaCliente
        ItemCategoriaCliente.ListIndex = -1
        ItemCategoriaCliente.Enabled = False

    End If

    Exit Sub

Erro_CategoriaCliente_Click:

    Select Case gErr

        Case 205422

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205423)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaCliente_Validate(Cancel As Boolean)

Dim lErro As Long, iCodigo As Integer

On Error GoTo Erro_CategoriaCliente_Validate

    If Len(CategoriaCliente.Text) <> 0 And CategoriaCliente.ListIndex = -1 Then

        'pesquisa a categoria na lista
        lErro = Combo_Item_Igual(CategoriaCliente)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 205424

        If lErro <> SUCESSO Then gError 205425

    End If
    
    'Se a Categoria estiver em Branco limpa e dasabilita
    If Len(CategoriaCliente.Text) = 0 Then
        ItemCategoriaCliente.Enabled = False
        ItemCategoriaCliente.Clear
    End If

    Exit Sub

Erro_CategoriaCliente_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 205424

        Case 205425
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", gErr, CategoriaCliente.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205426)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaProduto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CategoriaProduto_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colCategoria As New Collection

On Error GoTo Erro_CategoriaProduto_Click

    iAlterado = REGISTRO_ALTERADO

    If CategoriaProduto.ListIndex <> -1 Then

        'Preenche o objeto com a Categoria
         objCategoriaProduto.sCategoria = CategoriaProduto.Text

        'Lê os dados de itens de categorias de produto
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colCategoria)
        If lErro <> SUCESSO Then gError 205427
        
        ItemCategoriaProduto.Enabled = True
        ItemCategoriaProduto.Clear

        'Preenche ItemCategoriaProduto
        For Each objCategoriaProdutoItem In colCategoria

            ItemCategoriaProduto.AddItem (objCategoriaProdutoItem.sItem)

        Next

        TodosProdutos.Value = 0
    Else
        
        'Senão limpa e dasabilita o item
        ItemCategoriaProduto.ListIndex = -1
        ItemCategoriaProduto.Enabled = False
    End If

    Exit Sub

Erro_CategoriaProduto_Click:

    Select Case gErr

        Case 205427

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205428)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoExcPisCofins = Nothing
    Set colTiposTribPisCofins = Nothing

End Sub

Private Sub CategoriaProduto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_CategoriaProduto_Validate

    If Len(CategoriaProduto) <> 0 And CategoriaProduto.ListIndex = -1 Then

        'pesquisa a categoria na lista
        lErro = Combo_Item_Igual(CategoriaProduto)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 205429

        If lErro <> SUCESSO Then gError 205430

    End If

    If Len(CategoriaProduto) = 0 Then
        'Se item Categoria estiver em Branco limpa e dasabilita
        ItemCategoriaProduto.Enabled = False
        ItemCategoriaProduto.Clear
    End If

    Exit Sub

Erro_CategoriaProduto_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 205429

        Case 205430
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", gErr, CategoriaProduto.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205431)

    End Select

    Exit Sub

End Sub

Private Sub COFINSAliquota_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub COFINSAliquotaRS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub OptCliForn_Click(Index As Integer)
    
    If OptCliForn(PISCOFINSEXCECOES_TIPOCLIFORN_CLIENTE).Value = True Then
    
        ItemCategoriaFornecedor.Text = ""
        CategoriaFornecedor.Text = ""
        
        Frame3(PISCOFINSEXCECOES_TIPOCLIFORN_CLIENTE).Visible = True
        Frame3(PISCOFINSEXCECOES_TIPOCLIFORN_FORNECEDOR).Visible = False
    
    Else
        ItemCategoriaCliente.Text = ""
        CategoriaCliente.Text = ""
        
        Frame3(PISCOFINSEXCECOES_TIPOCLIFORN_CLIENTE).Visible = False
        Frame3(PISCOFINSEXCECOES_TIPOCLIFORN_FORNECEDOR).Visible = True
    
    End If

End Sub

Private Sub optPISCOFINS_Click(Index As Integer)
    iAlterado = REGISTRO_ALTERADO
    FramePIS.Enabled = True
    FrameCOFINS.Enabled = True
    If optPISCOFINS.Item(EXCECAO_PIS_COFINS_TIPO_PIS).Value Then
        FrameCOFINS.Enabled = False
        COFINSTipoTributacao.ListIndex = -1
        COFINSTipoCalculo.ListIndex = -1
        COFINSAliquota.Text = ""
        COFINSAliquotaRS.Text = ""
    ElseIf optPISCOFINS.Item(EXCECAO_PIS_COFINS_TIPO_COFINS).Value Then
        FramePIS.Enabled = False
        PISTipoTributacao.ListIndex = -1
        PISTipoCalculo.ListIndex = -1
        PISAliquota.Text = ""
        PISAliquotaRS.Text = ""
    End If
End Sub

Private Sub PISAliquota_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PISAliquotaRS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ItemCategoriaCliente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ItemCategoriaCliente_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ItemCategoriaFornecedor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ItemCategoriaFornecedor_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ItemCategoriaProduto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ItemCategoriaProduto_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LabelFundamentacao_Click()

Dim lErro As Long
Dim objExcPisCofins As New ClassPISCOFINSExcecao
Dim colSelecao As New Collection

On Error GoTo Erro_LabelFundamentacao_Click
    
    'Chama a LIsta de Excecoes de IPI
    Call Chama_Tela("ExcPisCofinsLista", colSelecao, objExcPisCofins, objEventoExcPisCofins)

    Exit Sub

Erro_LabelFundamentacao_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205432)

    End Select

    Exit Sub

End Sub

Private Sub objEventoExcPisCofins_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objExcPisCofins As ClassPISCOFINSExcecao

On Error GoTo Erro_objEventoExcPisCofins_evSelecao

    Set objExcPisCofins = obj1
    
    'Traz a Excecao para a Tela
    Call Traz_Excecao_Tela(objExcPisCofins)

    Me.Show

    Exit Sub

Erro_objEventoExcPisCofins_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205433)

    End Select

    Exit Sub

End Sub

Function Verifica_Identificacao_Preenchida() As Long

'verifica se todos os dados necessarios p/identificacao de uma excecao foram preenchidos
Dim lErro As Long

On Error GoTo Erro_Verifica_Identificacao_Preenchida

    If Len(Fundamentacao.Text) = 0 Then gError 205435

    'Testa se TodosProdutos está marcado
    If TodosProdutos.Value = 0 Then

        'Testa se Categoria do produto está preenchida
        If Len(CategoriaProduto.Text) = 0 Then gError 205436

        'Testa se Valor da Categoria do produto está preenchida
        If Len(ItemCategoriaProduto.Text) = 0 Then gError 205437

    End If
    
    If OptCliForn(PISCOFINSEXCECOES_TIPOCLIFORN_CLIENTE).Value = True Then
    
        'Testa se TodosProdutos está marcado
        If TodosClientes.Value = 0 Then
        
            'Testa se Categoria do cliente está preenchida
            If Len(CategoriaCliente.Text) = 0 Then gError 205438
        
            'Testa se Valor da Categoria do cliente está preenchida
            If Len(ItemCategoriaCliente.Text) = 0 Then gError 205439
    
        End If
    
    Else
    
        'Testa se TodosClientes está marcado
        If TodosFornecedores.Value = vbUnchecked Then
    
            'Testa se Categoria do cliente está preenchida
            If Len(CategoriaFornecedor.Text) = 0 Then gError 140405
    
            'Testa se Valor da Categoria do cliente está preenchida
            If Len(ItemCategoriaFornecedor.Text) = 0 Then gError 140406
    
        End If
    
    End If
    
    Verifica_Identificacao_Preenchida = SUCESSO

    Exit Function

Erro_Verifica_Identificacao_Preenchida:

     Verifica_Identificacao_Preenchida = gErr

     Select Case gErr

        Case 205435
            Call Rotina_Erro(vbOKOnly, "ERRO_FUNDAMENTACAO_NAO_PREENCHIDA", gErr)

        Case 205436
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_NAO_INFORMADA", gErr)

        Case 205437
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO1", gErr)

        Case 205438
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_CLIENTE_NAO_PREENCHIDA", gErr)

        Case 205439
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_CLIENTE_ITEM_NAO_PREENCHIDA", gErr)

        Case 140405
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_FORNECEDOR_NAO_PREENCHIDA", gErr)

        Case 140406
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_FORNECEDOR_ITEM_NAO_PREENCHIDA", gErr)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205440)

     End Select

     Exit Function

End Function

Private Function Move_Identificacao_Memoria(objPisCofinsExcecao As ClassPISCOFINSExcecao) As Long

    objPisCofinsExcecao.sFundamentacao = Fundamentacao.Text
    objPisCofinsExcecao.sCategoriaProduto = CategoriaProduto.Text
    objPisCofinsExcecao.sCategoriaProdutoItem = ItemCategoriaProduto.Text
    objPisCofinsExcecao.sCategoriaCliente = CategoriaCliente.Text
    objPisCofinsExcecao.sCategoriaClienteItem = ItemCategoriaCliente.Text
    objPisCofinsExcecao.sCategoriaFornecedor = CategoriaFornecedor.Text
    objPisCofinsExcecao.sCategoriaFornecedorItem = ItemCategoriaFornecedor.Text
    
    If OptCliForn(PISCOFINSEXCECOES_TIPOCLIFORN_CLIENTE).Value = True Then
    
        objPisCofinsExcecao.iTipoCliForn = PISCOFINSEXCECOES_TIPOCLIFORN_CLIENTE
    
    Else
    
        objPisCofinsExcecao.iTipoCliForn = PISCOFINSEXCECOES_TIPOCLIFORN_FORNECEDOR
    
    End If

End Function

Private Sub BotaoExcluir_Click()

Dim objPisCofinsExcecao As New ClassPISCOFINSExcecao
Dim colCategoria As New Collection
Dim colCategoriaItem As New Collection
Dim objCategoriaClienteItem As New ClassCategoriaClienteItem
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim lErro As Long
Dim vbMsgRet As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se todos os campos foram preenchidos
    If Verifica_Identificacao_Preenchida <> SUCESSO Then gError 205441

    'Pede Confirmação para exclusão ao usuário
    vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_EXCECAO_PIS_COFINS")

    If vbMsgRet = vbYes Then
        
        'Preenche o objPisCofinsExcecao
        If Move_Identificacao_Memoria(objPisCofinsExcecao) <> SUCESSO Then gError 205442
        
        'Exclui a Execeção IPI
        lErro = CF("PisCofinsExcecoes_Exclui", objPisCofinsExcecao)
        If lErro <> SUCESSO Then gError 205443
        
        'Limpa a tela
        Call Limpa_Tela_ExcPisCofins

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 205441 To 205443

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205444)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    'Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 205445
    
    'LImpa a tela
    Call Limpa_Tela_ExcPisCofins

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 205445 'Tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205446)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 205447
    
    'Limpa a tela
    Call Limpa_Tela_ExcPisCofins

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 205447 'Tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205448)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_ExcPisCofins()

    Call Limpa_Tela(Me)

    TodosProdutos.Value = vbUnchecked
    CategoriaProduto.ListIndex = -1
    TodosClientes.Value = vbUnchecked
    CategoriaCliente.ListIndex = -1
    PISTipoTributacao.ListIndex = -1
    COFINSTipoTributacao.ListIndex = -1
    PISTipoCalculo.ListIndex = -1
    COFINSTipoCalculo.ListIndex = -1
    PISTipoTributacaoE.ListIndex = -1
    COFINSTipoTributacaoE.ListIndex = -1
    
    TodosFornecedores.Value = vbUnchecked
    CategoriaFornecedor.ListIndex = -1
    
    optPISCOFINS.Item(EXCECAO_PIS_COFINS_TIPO_AMBOS).Value = True
    Call optPISCOFINS_Click(EXCECAO_PIS_COFINS_TIPO_AMBOS)
    
    Call CategoriaProduto_Click
    Call CategoriaCliente_Click
    Call CategoriaFornecedor_Click
    Call PISTipoTributacao_Click
    Call COFINSTipoTributacao_Click

    iAlterado = 0

    Exit Sub

End Sub

Private Sub CategoriaCliente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CategoriaFornecedor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim sCodigo As String
Dim iIndice As Integer
Dim objTipoTribPISCOFINS As New ClassTipoTribPISCOFINS
Dim colCategoriaProduto As New Collection
Dim colCategoriaCliente As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaCliente As New ClassCategoriaCliente
Dim colCategoriaFornecedor As New Collection
Dim objCategoriaFornecedor As New ClassCategoriaFornecedor

On Error GoTo Erro_Form_Load

    'Le as categorias de produto
    lErro = CF("CategoriasProduto_Le_Todas", colCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22542 Then gError 205449

    'Preenche CategoriaProduto
    For Each objCategoriaProduto In colCategoriaProduto

        CategoriaProduto.AddItem objCategoriaProduto.sCategoria

    Next

    'Le as categorias de cliente
    lErro = CF("CategoriaCliente_Le_Todos", colCategoriaCliente)
    If lErro <> SUCESSO Then gError 205450

    'Preenche CategoriaCliente
    For Each objCategoriaCliente In colCategoriaCliente
        CategoriaCliente.AddItem objCategoriaCliente.sCategoria
    Next

    'Le as categorias de Fornecedor
    lErro = CF("CategoriaFornecedor_Le_Todos", colCategoriaFornecedor)
    If lErro <> SUCESSO And lErro <> 68486 Then gError 205450

    'Preenche CategoriaFornecedor
    For Each objCategoriaFornecedor In colCategoriaFornecedor
        CategoriaFornecedor.AddItem objCategoriaFornecedor.sCategoria
    Next
    
    'Tipo de PIS e Tipo de COFINS
    lErro = CF("TiposTribPISCOFINS_Le_Todos", colTiposTribPisCofins)
    If lErro <> SUCESSO Then gError 205451
    
    PISTipoTributacao.Clear
    COFINSTipoTributacao.Clear
    PISTipoTributacaoE.Clear
    COFINSTipoTributacaoE.Clear
    
    For Each objTipoTribPISCOFINS In colTiposTribPisCofins
        sCodigo = CStr(objTipoTribPISCOFINS.iTipo) & SEPARADOR & objTipoTribPISCOFINS.sDescricao
        PISTipoTributacao.AddItem (sCodigo)
        PISTipoTributacao.ItemData(PISTipoTributacao.NewIndex) = objTipoTribPISCOFINS.iTipo
        COFINSTipoTributacao.AddItem (sCodigo)
        COFINSTipoTributacao.ItemData(PISTipoTributacao.NewIndex) = objTipoTribPISCOFINS.iTipo
        If objTipoTribPISCOFINS.iTipo >= 50 And objTipoTribPISCOFINS.iTipo <= 98 Then
            PISTipoTributacaoE.AddItem (sCodigo)
            PISTipoTributacaoE.ItemData(PISTipoTributacaoE.NewIndex) = objTipoTribPISCOFINS.iTipo
            COFINSTipoTributacaoE.AddItem (sCodigo)
            COFINSTipoTributacaoE.ItemData(PISTipoTributacaoE.NewIndex) = objTipoTribPISCOFINS.iTipo
        End If
    Next

    Set objEventoExcPisCofins = New AdmEvento
    
    'Tipo de Cálculo do PIS
    PISTipoCalculo.Clear

    PISTipoCalculo.AddItem TRIB_TIPO_CALCULO_VALOR & SEPARADOR & TRIB_TIPO_CALCULO_VALOR_TEXTO
    PISTipoCalculo.ItemData(PISTipoCalculo.NewIndex) = TRIB_TIPO_CALCULO_VALOR

    PISTipoCalculo.AddItem TRIB_TIPO_CALCULO_PERCENTUAL & SEPARADOR & TRIB_TIPO_CALCULO_PERCENTUAL_TEXTO
    PISTipoCalculo.ItemData(PISTipoCalculo.NewIndex) = TRIB_TIPO_CALCULO_PERCENTUAL

    'Tipo de Cálculo do COFINS
    COFINSTipoCalculo.Clear

    COFINSTipoCalculo.AddItem TRIB_TIPO_CALCULO_VALOR & SEPARADOR & TRIB_TIPO_CALCULO_VALOR_TEXTO
    COFINSTipoCalculo.ItemData(COFINSTipoCalculo.NewIndex) = TRIB_TIPO_CALCULO_VALOR

    COFINSTipoCalculo.AddItem TRIB_TIPO_CALCULO_PERCENTUAL & SEPARADOR & TRIB_TIPO_CALCULO_PERCENTUAL_TEXTO
    COFINSTipoCalculo.ItemData(COFINSTipoCalculo.NewIndex) = TRIB_TIPO_CALCULO_PERCENTUAL

    Call CategoriaProduto_Click
    Call CategoriaCliente_Click
    Call PISTipoTributacao_Click
    Call COFINSTipoTributacao_Click
    Call CategoriaFornecedor_Click

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 205449 To 205451

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205452)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objExcPisCofins As ClassPISCOFINSExcecao) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se objIPIExcecao estiver preenchido
    If Not (objExcPisCofins Is Nothing) Then
        
        'Preenche a tela
        Call Traz_Excecao_Tela(objExcPisCofins)

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205453)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Function Gravar_Registro()

Dim lErro As Long
Dim iIndice As Integer
Dim objPisCofinsExcecao As New ClassPISCOFINSExcecao
Dim objCategoriaClienteItem As ClassCategoriaClienteItem
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem
Dim colCategoriaItem As New Collection
Dim objTipoTrib As ClassTipoTribPISCOFINS

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se todos os campos foram preenchidos corretamente
    If Verifica_Identificacao_Preenchida <> SUCESSO Then gError 205454
    
    'Passa os dados para objPisCofinsExcecao
    If Move_Identificacao_Memoria(objPisCofinsExcecao) <> SUCESSO Then gError 205455

    If optPISCOFINS.Item(EXCECAO_PIS_COFINS_TIPO_AMBOS).Value Or optPISCOFINS.Item(EXCECAO_PIS_COFINS_TIPO_PIS).Value Then
        If PISTipoTributacao.ListIndex = -1 Then gError 205456
        objPisCofinsExcecao.iTipoPIS = PISTipoTributacao.ItemData(PISTipoTributacao.ListIndex)
        If PISTipoTributacaoE.ListIndex <> -1 Then objPisCofinsExcecao.iTipoPISE = PISTipoTributacaoE.ItemData(PISTipoTributacaoE.ListIndex)
    End If
    
    If optPISCOFINS.Item(EXCECAO_PIS_COFINS_TIPO_AMBOS).Value Or optPISCOFINS.Item(EXCECAO_PIS_COFINS_TIPO_COFINS).Value Then
        If COFINSTipoTributacao.ListIndex = -1 Then gError 205457
        objPisCofinsExcecao.iTipoCOFINS = COFINSTipoTributacao.ItemData(COFINSTipoTributacao.ListIndex)
        If COFINSTipoTributacaoE.ListIndex <> -1 Then objPisCofinsExcecao.iTipoCOFINSE = COFINSTipoTributacaoE.ItemData(COFINSTipoTributacaoE.ListIndex)
    End If
    
    objPisCofinsExcecao.iPISTipoCalculo = Codigo_Extrai(PISTipoCalculo.Text)
    objPisCofinsExcecao.iCOFINSTipoCalculo = Codigo_Extrai(COFINSTipoCalculo.Text)
    
    If objPisCofinsExcecao.iTipoPISE <> 0 Then
        For Each objTipoTrib In colTiposTribPisCofins
            If objTipoTrib.iTipo = objPisCofinsExcecao.iTipoPISE Then
                Select Case objPisCofinsExcecao.iPISTipoCalculo
                    Case TIPO_TRIB_TIPO_CALCULO_PERCENTUAL
                        If objTipoTrib.iTipoCalculo = TIPO_TRIB_TIPO_CALCULO_VALOR Or objTipoTrib.iTipoCalculo = TIPO_TRIB_TIPO_CALCULO_DESABILITADO Then gError 211316
                    Case TIPO_TRIB_TIPO_CALCULO_VALOR
                        If objTipoTrib.iTipoCalculo = TIPO_TRIB_TIPO_CALCULO_PERCENTUAL Or objTipoTrib.iTipoCalculo = TIPO_TRIB_TIPO_CALCULO_DESABILITADO Then gError 211317
                End Select
                Exit For
            End If
        Next
    End If
    
    If objPisCofinsExcecao.iTipoCOFINSE <> 0 Then
        For Each objTipoTrib In colTiposTribPisCofins
            If objTipoTrib.iTipo = objPisCofinsExcecao.iTipoCOFINSE Then
                Select Case objPisCofinsExcecao.iCOFINSTipoCalculo
                    Case TIPO_TRIB_TIPO_CALCULO_PERCENTUAL
                        If objTipoTrib.iTipoCalculo = TIPO_TRIB_TIPO_CALCULO_VALOR Or objTipoTrib.iTipoCalculo = TIPO_TRIB_TIPO_CALCULO_DESABILITADO Then gError 211318
                    Case TIPO_TRIB_TIPO_CALCULO_VALOR
                        If objTipoTrib.iTipoCalculo = TIPO_TRIB_TIPO_CALCULO_PERCENTUAL Or objTipoTrib.iTipoCalculo = TIPO_TRIB_TIPO_CALCULO_DESABILITADO Then gError 211319
                End Select
                Exit For
            End If
        Next
    End If

    'Se campos habilitados, move seus dados
    If PISAliquota.Enabled Then
        If Len(PISAliquota.Text) > 0 Then objPisCofinsExcecao.dAliquotaPisPerc = CDbl(PISAliquota.Text / 100)
    End If
    If PISAliquotaRS.Enabled Then
        If Len(PISAliquotaRS.Text) > 0 Then objPisCofinsExcecao.dAliquotaPisRS = StrParaDbl(PISAliquotaRS.Text)
    End If
    
    If COFINSAliquota.Enabled Then
        If Len(COFINSAliquota.Text) > 0 Then objPisCofinsExcecao.dAliquotaCofinsPerc = CDbl(COFINSAliquota.Text / 100)
    End If
    If COFINSAliquotaRS.Enabled Then
        If Len(COFINSAliquotaRS.Text) > 0 Then objPisCofinsExcecao.dAliquotaCofinsRS = StrParaDbl(COFINSAliquotaRS.Text)
    End If
    
    If optPISCOFINS.Item(EXCECAO_PIS_COFINS_TIPO_COFINS).Value Then
        objPisCofinsExcecao.iTipo = EXCECAO_PIS_COFINS_TIPO_COFINS
    ElseIf optPISCOFINS.Item(EXCECAO_PIS_COFINS_TIPO_PIS).Value Then
        objPisCofinsExcecao.iTipo = EXCECAO_PIS_COFINS_TIPO_PIS
    Else
        objPisCofinsExcecao.iTipo = EXCECAO_PIS_COFINS_TIPO_AMBOS
    End If
    
    'identifica o tipo da prioridade
    If (TodosClientes.Value = 1 And objPisCofinsExcecao.iTipoCliForn = PISCOFINSEXCECOES_TIPOCLIFORN_CLIENTE) _
        Or (TodosFornecedores.Value = 1 And objPisCofinsExcecao.iTipoCliForn = PISCOFINSEXCECOES_TIPOCLIFORN_FORNECEDOR) Then
        objPisCofinsExcecao.iPrioridade = TIPOTRIB_PRIORIDADE_PRODUTO
    Else
        If TodosProdutos.Value = 1 Then
            objPisCofinsExcecao.iPrioridade = TIPOTRIB_PRIORIDADE_CLIENTE
        Else
            objPisCofinsExcecao.iPrioridade = TIPOTRIB_PRIORIDADE_CLIENTE_PRODUTO
        End If
    End If
    
    lErro = Trata_Alteracao(objPisCofinsExcecao, objPisCofinsExcecao.sCategoriaCliente, objPisCofinsExcecao.sCategoriaClienteItem, objPisCofinsExcecao.sCategoriaProduto, objPisCofinsExcecao.sCategoriaProdutoItem)
    If lErro <> SUCESSO Then gError 205458
    
    lErro = CF("PisCofinsExcecoes_Grava", objPisCofinsExcecao)
    If lErro <> SUCESSO Then gError 205459

    Call Limpa_Tela_ExcPisCofins

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 205454, 205455, 205458, 205459

        Case 205456, 205457
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_NAO_PREENCHIDO", gErr)

        Case 211316, 211317
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_CALC_PIS_INCOMPATIVEL", gErr)
        
        Case 211318, 211319
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_CALC_COFINS_INCOMPATIVEL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205460)

    End Select

    Exit Function

End Function

Private Sub Fundamentacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ItemCategoriaCliente_Validate(Cancel As Boolean)

Dim lErro As Long, iCodigo As Integer
On Error GoTo Erro_ItemCategoriaCliente_Validate

    If Len(ItemCategoriaCliente.Text) <> 0 And ItemCategoriaCliente.ListIndex = -1 Then

        'pesquisa o item na lista
        lErro = Combo_Item_Igual(ItemCategoriaCliente)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 205461

        If lErro <> SUCESSO Then gError 205462

    End If

    Exit Sub

Erro_ItemCategoriaCliente_Validate:

    Cancel = True


    Select Case gErr

        Case 205461

        Case 205462
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTEITEM_INEXISTENTE", gErr, ItemCategoriaCliente.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205463)

    End Select

    Exit Sub

End Sub

Private Sub ItemCategoriaProduto_Validate(Cancel As Boolean)

Dim lErro As Long, iCodigo As Integer

On Error GoTo Erro_ItemCategoriaProduto_Validate

    If Len(ItemCategoriaProduto.Text) <> 0 And ItemCategoriaProduto.ListIndex = -1 Then

        'pesquisa o item na lista
        lErro = Combo_Item_Igual(ItemCategoriaProduto)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 205464

        If lErro <> SUCESSO Then gError 205465

    End If

    Exit Sub

Erro_ItemCategoriaProduto_Validate:

    Cancel = True

    Select Case gErr

        Case 205464 'Tratado na rotina chamada

        Case 205465
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", gErr, ItemCategoriaProduto.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205466)

    End Select

    Exit Sub

End Sub

Private Sub ComboPISCOFINSTipo_Click(ByVal iTipo As Integer, ByVal sImposto As String)

Dim objTipoTribPISCOFINS As ClassTipoTribPISCOFINS

On Error GoTo Erro_ComboPISCOFINSTipo_Click

    For Each objTipoTribPISCOFINS In colTiposTribPisCofins
        If objTipoTribPISCOFINS.iTipo = iTipo Then Exit For
    Next
    If Not (objTipoTribPISCOFINS Is Nothing) Then
    
        Select Case objTipoTribPISCOFINS.iTipoCalculo
            
            Case TIPO_TRIB_TIPO_CALCULO_DESABILITADO
                Call TipoCalculo_Click(-1, sImposto)
                Controls(sImposto & "TipoCalculo").Enabled = False
    
            Case TIPO_TRIB_TIPO_CALCULO_PERCENTUAL
                Call Combo_Seleciona_ItemData(Controls(sImposto & "TipoCalculo"), TIPO_TRIB_TIPO_CALCULO_PERCENTUAL)
                Call TipoCalculo_Click(TRIB_TIPO_CALCULO_PERCENTUAL, sImposto)
                Controls(sImposto & "TipoCalculo").Enabled = False
    
            Case TIPO_TRIB_TIPO_CALCULO_VALOR
                Call Combo_Seleciona_ItemData(Controls(sImposto & "TipoCalculo"), TRIB_TIPO_CALCULO_VALOR)
                Call TipoCalculo_Click(TRIB_TIPO_CALCULO_VALOR, sImposto)
                Controls(sImposto & "TipoCalculo").Enabled = False
        
            Case TIPO_TRIB_TIPO_CALCULO_ESCOLHA
                Call TipoCalculo_Click(-2, sImposto)
                Controls(sImposto & "TipoCalculo").Enabled = True
        
        End Select
        
    End If
        
    Exit Sub

Erro_ComboPISCOFINSTipo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205467)

    End Select

    Exit Sub

End Sub

Private Sub TodosProdutos_Click()

Dim lErro As Long
On Error GoTo Erro_TodosProdutos_Click

    'TodosCLientes e todos Produto não podem ser marcados ao mesmo tempo
    If TodosClientes.Value = vbChecked And TodosProdutos.Value = vbChecked Then gError 205468
    
    'If TodosProdutos.Value = 1 And TodosProdutos.Value = 1 Then gError 21504
    If TodosProdutos.Value = 1 Then CategoriaProduto.ListIndex = -1

    Exit Sub

Erro_TodosProdutos_Click:

    Select Case gErr

        Case 205468
            Call Rotina_Erro(vbOKOnly, "AVISO_NAO_E_POSSIVEL_SELECIONAR_TODOS", gErr)
            TodosProdutos.Value = vbUnchecked
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205469)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_EXCECOES_IPI
    Set Form_Load_Ocx = Me
    Caption = "Exceções de Pis e Cofins"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ExcPisCofins"
    
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
        
        If Me.ActiveControl Is Fundamentacao Then
            Call LabelFundamentacao_Click
        End If
    
    End If

End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub LabelFundamentacao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFundamentacao, Source, X, Y)
End Sub

Private Sub LabelFundamentacao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFundamentacao, Button, Shift, X, Y)
End Sub

Private Sub TodosClientes_Click()

Dim lErro As Long
On Error GoTo Erro_TodosClientes_Click

    'TodosCLientes e todos Produto não podem ser marcados ao mesmo tempo
    If TodosProdutos.Value = 1 And TodosClientes.Value = 1 And OptCliForn(PISCOFINSEXCECOES_TIPOCLIFORN_CLIENTE).Value = True Then gError 205470

    If TodosClientes.Value = 1 Then CategoriaCliente.ListIndex = -1

    Exit Sub

Erro_TodosClientes_Click:

    Select Case gErr

        Case 205470
            Call Rotina_Erro(vbOKOnly, "AVISO_NAO_E_POSSIVEL_SELECIONAR_TODOS", gErr)
            TodosClientes.Value = 0

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205471)

    End Select

    Exit Sub

End Sub

Private Sub PISTipoCalculo_Click()
    Call TipoCalculo_Click(Codigo_Extrai(PISTipoCalculo.Text), "PIS")
End Sub

Private Sub COFINSTipoCalculo_Click()
    Call TipoCalculo_Click(Codigo_Extrai(COFINSTipoCalculo.Text), "COFINS")
End Sub

Private Sub TipoCalculo_Click(ByVal iTipo As Integer, ByVal sImposto As String, Optional ByVal bAtualizaTrib As Boolean = True)

On Error GoTo Erro_TipoCalculo_Click

    '-2 = Respeita o que estã no tipo
    If iTipo = -2 Then
        If Len(Trim(Controls(sImposto & "TipoCalculo"))) > 0 Then
            iTipo = Codigo_Extrai(Controls(sImposto & "TipoCalculo").Text)
        Else
            iTipo = -1
        End If
    End If

    Select Case iTipo
   
        Case -1
            Controls(sImposto & "TipoCalculo").ListIndex = -1
            Controls(sImposto & "AliquotaRS").Enabled = False
            Controls(sImposto & "Aliquota").Enabled = False
            Controls(sImposto & "AliquotaRS").Text = ""
            Controls(sImposto & "Aliquota").Text = ""
        
        Case TRIB_TIPO_CALCULO_VALOR
            Controls(sImposto & "AliquotaRS").Enabled = True
            Controls(sImposto & "Aliquota").Enabled = False
            Controls(sImposto & "Aliquota").Text = ""
        
        Case TRIB_TIPO_CALCULO_PERCENTUAL
            Controls(sImposto & "AliquotaRS").Enabled = False
            Controls(sImposto & "Aliquota").Enabled = True
            Controls(sImposto & "AliquotaRS").Text = ""
    
    End Select
        
    Exit Sub

Erro_TipoCalculo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205472)

    End Select

    Exit Sub
    
End Sub

Public Sub PISTipoTributacao_Click()
    Call ComboPISCOFINSTipo_Click(Codigo_Extrai(PISTipoTributacao.Text), "PIS")
End Sub

Public Sub COFINSTipoTributacao_Click()
    Call ComboPISCOFINSTipo_Click(Codigo_Extrai(COFINSTipoTributacao.Text), "COFINS")
End Sub

Private Sub COFINSAliquota_Validate(Cancel As Boolean)
    Call Aliquota_Validate(COFINSAliquota, Cancel)
End Sub

Private Sub COFINSAliquotaRS_Validate(Cancel As Boolean)
    Call Valor_Validate(COFINSAliquotaRS, Cancel)
End Sub

Private Sub PISAliquota_Validate(Cancel As Boolean)
    Call Aliquota_Validate(PISAliquota, Cancel)
End Sub

Private Sub PISAliquotaRS_Validate(Cancel As Boolean)
    Call Valor_Validate(PISAliquotaRS, Cancel)
End Sub

Private Sub CategoriaFornecedor_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objCategoriaFornecedorItem As New ClassCategoriaFornItem
Dim objCategoriaFornecedor As New ClassCategoriaFornecedor
Dim colCategoria As New Collection

On Error GoTo Erro_CategoriaFornecedor_Click

    iAlterado = REGISTRO_ALTERADO

    'Verifica se a CategoriaFornecedor foi preenchida
    If CategoriaFornecedor.ListIndex <> -1 Then

        objCategoriaFornecedorItem.sCategoria = CategoriaFornecedor.Text

        'Lê os dados de Itens da Categoria do Fornecedor
        lErro = CF("CategoriaFornecedor_Le_Itens", objCategoriaFornecedorItem, colCategoria)
        If lErro <> SUCESSO Then gError 205422

        ItemCategoriaFornecedor.Enabled = True

        'Limpa os dados de ItemCategoriaFornecedor
        ItemCategoriaFornecedor.Clear

        'Preenche ItemCategoriaFornecedor
        For Each objCategoriaFornecedorItem In colCategoria

            ItemCategoriaFornecedor.AddItem objCategoriaFornecedorItem.sItem

        Next
        TodosFornecedores.Value = 0

    Else
        
        'Senão Desabilita e limpa ItemCategoriaFornecedor
        ItemCategoriaFornecedor.ListIndex = -1
        ItemCategoriaFornecedor.Enabled = False

    End If

    Exit Sub

Erro_CategoriaFornecedor_Click:

    Select Case gErr

        Case 205422

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205423)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaFornecedor_Validate(Cancel As Boolean)

Dim lErro As Long, iCodigo As Integer

On Error GoTo Erro_CategoriaFornecedor_Validate

    If Len(CategoriaFornecedor.Text) <> 0 And CategoriaFornecedor.ListIndex = -1 Then

        'pesquisa a categoria na lista
        lErro = Combo_Item_Igual(CategoriaFornecedor)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 205424

        If lErro <> SUCESSO Then gError 205425

    End If
    
    'Se a Categoria estiver em Branco limpa e dasabilita
    If Len(CategoriaFornecedor.Text) = 0 Then
        ItemCategoriaFornecedor.Enabled = False
        ItemCategoriaFornecedor.Clear
    End If

    Exit Sub

Erro_CategoriaFornecedor_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 205424

        Case 205425
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", gErr, CategoriaFornecedor.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205426)

    End Select

    Exit Sub

End Sub

Private Sub ItemCategoriaFornecedor_Validate(Cancel As Boolean)

Dim lErro As Long, iCodigo As Integer
On Error GoTo Erro_ItemCategoriaFornecedor_Validate

    If Len(ItemCategoriaFornecedor.Text) <> 0 And ItemCategoriaFornecedor.ListIndex = -1 Then

        'pesquisa o item na lista
        lErro = Combo_Item_Igual(ItemCategoriaFornecedor)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 205461

        If lErro <> SUCESSO Then gError 205462

    End If

    Exit Sub

Erro_ItemCategoriaFornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 205461

        Case 205462
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAFORNECEDORITEM_INEXISTENTE", gErr, ItemCategoriaFornecedor.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205463)

    End Select

    Exit Sub

End Sub

Private Sub TodosFornecedores_Click()

Dim lErro As Long
On Error GoTo Erro_TodosFornecedores_Click

    'Todos Fornecedores e todos Produto não podem ser marcados ao mesmo tempo
    If TodosProdutos.Value = 1 And TodosFornecedores.Value = 1 And OptCliForn(PISCOFINSEXCECOES_TIPOCLIFORN_FORNECEDOR).Value = True Then gError 205470

    If TodosFornecedores.Value = 1 Then CategoriaFornecedor.ListIndex = -1

    Exit Sub

Erro_TodosFornecedores_Click:

    Select Case gErr

        Case 205470
            Call Rotina_Erro(vbOKOnly, "AVISO_NAO_E_POSSIVEL_SELECIONAR_TODOS", gErr)
            TodosFornecedores.Value = 0

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205471)

    End Select

    Exit Sub

End Sub


