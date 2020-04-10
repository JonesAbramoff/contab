VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl MatPrimOcx 
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8190
   ScaleHeight     =   7095
   ScaleWidth      =   8190
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6105
      Index           =   1
      Left            =   150
      TabIndex        =   4
      Top             =   780
      Width           =   7875
      Begin VB.CommandButton BotaoCustoMP 
         Caption         =   "Custo de Matéria-Prima"
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
         Left            =   2775
         TabIndex        =   34
         Top             =   5565
         Width           =   2340
      End
      Begin VB.CommandButton BotaoDetalhar 
         Caption         =   "Análise de Custo de Produto Intermediário"
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
         Left            =   5235
         TabIndex        =   22
         Top             =   5565
         Width           =   2535
      End
      Begin VB.Frame Frame2 
         Caption         =   "Insumos"
         Height          =   3825
         Left            =   75
         TabIndex        =   15
         Top             =   1620
         Width           =   7710
         Begin VB.TextBox DescricaoIns 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1335
            MaxLength       =   50
            TabIndex        =   20
            Top             =   705
            Width           =   1965
         End
         Begin VB.TextBox PercentualParticipacaoIns 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   3015
            MaxLength       =   50
            TabIndex        =   17
            Top             =   660
            Width           =   1170
         End
         Begin MSMask.MaskEdBox CustoConsiderarIns 
            Height          =   225
            Left            =   5535
            TabIndex        =   16
            Top             =   690
            Visible         =   0   'False
            Width           =   1245
            _ExtentX        =   2196
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CustoUnitarioIns 
            Height          =   225
            Left            =   4305
            TabIndex        =   18
            Top             =   720
            Visible         =   0   'False
            Width           =   1155
            _ExtentX        =   2037
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoIns 
            Height          =   225
            Left            =   315
            TabIndex        =   19
            Top             =   660
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridInsumo 
            Height          =   2205
            Left            =   105
            TabIndex        =   21
            Top             =   270
            Width           =   7020
            _ExtentX        =   12383
            _ExtentY        =   3889
            _Version        =   393216
            Rows            =   6
            Cols            =   6
         End
         Begin VB.Label LabelTotalPartic 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   3480
            TabIndex        =   41
            Top             =   2625
            Width           =   1320
         End
         Begin VB.Label Label3 
            Caption         =   "Total de Participação:"
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
            Left            =   1455
            TabIndex        =   40
            Top             =   2670
            Width           =   1935
         End
         Begin VB.Label Label5 
            Caption         =   "Renda:"
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
            Left            =   5235
            TabIndex        =   39
            Top             =   3060
            Width           =   600
         End
         Begin VB.Label LabelPerda 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   5895
            TabIndex        =   38
            Top             =   3015
            Width           =   1320
         End
         Begin VB.Label Label4 
            Caption         =   "Subtotal:"
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
            Left            =   5055
            TabIndex        =   37
            Top             =   2655
            Width           =   765
         End
         Begin VB.Label LabelSubtotalIns 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   5895
            TabIndex        =   36
            Top             =   2610
            Width           =   1320
         End
         Begin VB.Label LabelTotalIns 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   5895
            TabIndex        =   31
            Top             =   3420
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Total:"
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
            Left            =   5310
            TabIndex        =   30
            Top             =   3495
            Width           =   555
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Identificação"
         Height          =   1620
         Left            =   75
         TabIndex        =   5
         Top             =   15
         Width           =   7710
         Begin MSMask.MaskEdBox Produto 
            Height          =   315
            Left            =   960
            TabIndex        =   6
            Top             =   240
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            PromptChar      =   " "
         End
         Begin VB.Label ProdutoLbl 
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
            Left            =   195
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   14
            Top             =   285
            Width           =   735
         End
         Begin VB.Label LabelVersao 
            AutoSize        =   -1  'True
            Caption         =   "Versão:"
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
            TabIndex        =   13
            Top             =   750
            Width           =   660
         End
         Begin VB.Label Label7 
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
            Left            =   2640
            TabIndex        =   12
            Top             =   750
            Width           =   480
         End
         Begin VB.Label Versao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   960
            TabIndex        =   11
            Top             =   705
            Width           =   1500
         End
         Begin VB.Label Descricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2550
            TabIndex        =   10
            Top             =   240
            Width           =   5010
         End
         Begin VB.Label Data 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3180
            TabIndex        =   9
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label LabelObeservações 
            AutoSize        =   -1  'True
            Caption         =   "Obs:"
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
            TabIndex        =   8
            Top             =   1200
            Width           =   405
         End
         Begin VB.Label Observacao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   960
            TabIndex        =   7
            Top             =   1155
            Width           =   6570
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6015
      Index           =   2
      Left            =   135
      TabIndex        =   23
      Top             =   810
      Visible         =   0   'False
      Width           =   7875
      Begin VB.CommandButton BotaoCustoEmb 
         Caption         =   "Custo de Embalagem"
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
         Left            =   2685
         TabIndex        =   35
         Top             =   5505
         Width           =   2340
      End
      Begin MSMask.MaskEdBox CustoConsiderarEmb 
         Height          =   225
         Left            =   5955
         TabIndex        =   29
         Top             =   2430
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox QuantidadeEmb 
         Height          =   225
         Left            =   3930
         TabIndex        =   28
         Top             =   1950
         Width           =   990
         _ExtentX        =   1746
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CustoUnitarioEmb 
         Height          =   225
         Left            =   4740
         TabIndex        =   27
         Top             =   1980
         Visible         =   0   'False
         Width           =   1080
         _ExtentX        =   1905
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoEmb 
         Height          =   225
         Left            =   135
         TabIndex        =   26
         Top             =   1905
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.TextBox DescricaoEmb 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   1245
         MaxLength       =   50
         TabIndex        =   25
         Top             =   1920
         Width           =   2370
      End
      Begin MSFlexGridLib.MSFlexGrid GridEmbalagem 
         Height          =   4890
         Left            =   45
         TabIndex        =   24
         Top             =   285
         Width           =   7710
         _ExtentX        =   13600
         _ExtentY        =   8625
         _Version        =   393216
         Rows            =   6
         Cols            =   6
      End
      Begin VB.Label LabelTotalEmb 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   5745
         TabIndex        =   33
         Top             =   5520
         Width           =   1320
      End
      Begin VB.Label Label2 
         Caption         =   "Total:"
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
         Left            =   5145
         TabIndex        =   32
         Top             =   5595
         Width           =   555
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6855
      ScaleHeight     =   495
      ScaleWidth      =   1125
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1185
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   615
         Picture         =   "MatPrim.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   90
         Picture         =   "MatPrim.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6510
      Left            =   75
      TabIndex        =   3
      Top             =   465
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   11483
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Insumos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Embalagens"
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
Attribute VB_Name = "MatPrimOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Grid_Insumo
Dim objGridInsumo As AdmGrid
Dim iGrid_ProdutoIns_Col As Integer
Dim iGrid_DescricaoIns_Col As Integer
Dim iGrid_ParticipacaoIns_Col As Integer
Dim iGrid_CustoUnitarioIns_Col As Integer
Dim iGrid_CustoConsiderarIns_Col As Integer

'Grid_Embalagem
Dim objGridEmbalagem As AdmGrid
Dim iGrid_ProdutoEmb_Col As Integer
Dim iGrid_DescricaoEmb_Col As Integer
Dim iGrid_QuantidadeEmb_Col As Integer
Dim iGrid_CustoUnitarioEmb_Col As Integer
Dim iGrid_CustoConsiderarEmb_Col As Integer

'Numero maximo de linhas do grid
Const NUM_GRID_INSUMO = 100
Const NUM_GRID_EMBALAGEM = 100

'Browse
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

'variaveis de controle de alteração
Dim colComponentes As New Collection
Dim iFrameAtual As Integer
Public iAlterado As Integer

Private Function Inicializa_Grid_Insumo(objGridIns As AdmGrid) As Long
'realiza a inicialização do grid

    'tela em questão
    Set objGridIns.objForm = Me

    'titulos do grid
    objGridIns.colColuna.Add (" ")
    objGridIns.colColuna.Add ("Produto")
    objGridIns.colColuna.Add ("Descrição")
    objGridIns.colColuna.Add ("Participação")
    objGridIns.colColuna.Add ("Custo Unitário")
    objGridIns.colColuna.Add ("A Considerar")

   'campos de edição do grid
    objGridIns.colCampo.Add (ProdutoIns.Name)
    objGridIns.colCampo.Add (DescricaoIns.Name)
    objGridIns.colCampo.Add (PercentualParticipacaoIns.Name)
    objGridIns.colCampo.Add (CustoUnitarioIns.Name)
    objGridIns.colCampo.Add (CustoConsiderarIns.Name)

    'atribui valor as colunas
    iGrid_ProdutoIns_Col = 1
    iGrid_DescricaoIns_Col = 2
    iGrid_ParticipacaoIns_Col = 3
    iGrid_CustoUnitarioIns_Col = 4
    iGrid_CustoConsiderarIns_Col = 5
    
    objGridIns.objGrid = GridInsumo

    'todas as linhas do grid
    objGridIns.objGrid.Rows = NUM_GRID_INSUMO + 1
    
    'linhas visiveis do grid
    objGridIns.iLinhasVisiveis = 8

    GridInsumo.ColWidth(0) = 400

    objGridIns.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    Call Grid_Inicializa(objGridIns)
    
    Inicializa_Grid_Insumo = SUCESSO

End Function

Private Function Inicializa_Grid_Embalagem(objGridEmb As AdmGrid) As Long
'realiza a inicialização do grid

    'tela em questão
    Set objGridEmb.objForm = Me

    'titulos do grid
    objGridEmb.colColuna.Add (" ")
    objGridEmb.colColuna.Add ("Produto")
    objGridEmb.colColuna.Add ("Descrição")
    objGridEmb.colColuna.Add ("Quantidade")
    objGridEmb.colColuna.Add ("Custo Unitário")
    objGridEmb.colColuna.Add ("A Considerar")

   'campos de edição do grid
    objGridEmb.colCampo.Add (ProdutoEmb.Name)
    objGridEmb.colCampo.Add (DescricaoEmb.Name)
    objGridEmb.colCampo.Add (QuantidadeEmb.Name)
    objGridEmb.colCampo.Add (CustoUnitarioEmb.Name)
    objGridEmb.colCampo.Add (CustoConsiderarEmb.Name)

    'atribui valores as colunas
    iGrid_ProdutoEmb_Col = 1
    iGrid_DescricaoEmb_Col = 2
    iGrid_QuantidadeEmb_Col = 3
    iGrid_CustoUnitarioEmb_Col = 4
    iGrid_CustoConsiderarEmb_Col = 5
    
    objGridEmb.objGrid = GridEmbalagem

    'todas as linhas do grid
    objGridEmb.objGrid.Rows = NUM_GRID_EMBALAGEM + 1

    'linhas visiveis do grid
    objGridEmb.iLinhasVisiveis = 20

    GridEmbalagem.ColWidth(0) = 400

    objGridEmb.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    Call Grid_Inicializa(objGridEmb)
    
    Inicializa_Grid_Embalagem = SUCESSO

End Function

Function Saida_Celula_Insumo(objGridIns As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Insumo

    lErro = Grid_Inicializa_Saida_Celula(objGridIns)
    
    If lErro = SUCESSO Then
        
        'Verifica qual a coluna atual do Grid
        Select Case objGridIns.objGrid.Col
        
            'Coluna Produto
            Case iGrid_ProdutoIns_Col
                lErro = Saida_Celula_ProdutoIns(objGridIns)
                If lErro <> SUCESSO Then gError 123047
                
            'Coluna Descrição
            Case iGrid_DescricaoIns_Col
                lErro = Saida_Celula_DescricaoIns(objGridIns)
                If lErro <> SUCESSO Then gError 123048
                
            'Coluna Participacao
            Case iGrid_ParticipacaoIns_Col
                lErro = Saida_Celula_ParticipacaoIns(objGridIns)
                If lErro <> SUCESSO Then gError 123049
                
            'Coluna Custo Unitario
            Case iGrid_CustoUnitarioIns_Col
                lErro = Saida_Celula_CustoUnitarioIns(objGridIns)
                If lErro <> SUCESSO Then gError 123050
                
            'Coluna Custo Considerar
            Case iGrid_CustoConsiderarIns_Col
                lErro = Saida_Celula_CustoConsiderarIns(objGridIns)
                If lErro <> SUCESSO Then gError 123051
                
        End Select
        
        lErro = Grid_Finaliza_Saida_Celula(objGridIns)
        If lErro <> SUCESSO Then gError 123052

    End If
    
    Saida_Celula_Insumo = SUCESSO

    Exit Function

Erro_Saida_Celula_Insumo:

    Saida_Celula_Insumo = gErr

    Select Case gErr

        Case 123047, 123048, 123049, 123050, 123051, 123052

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162673)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ProdutoIns(objGridIns As AdmGrid) As Long
'Faz a critica da celula Produto do grid

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ProdutoIns

    Set objGridIns.objControle = ProdutoIns
    
    'Verifica se o Produto está vazio
    If Len(Trim(ProdutoIns.ClipText)) <> 0 Then
    
        'Realiza a soma de uma linha no grid
        If GridInsumo.Row - GridInsumo.FixedRows = objGridIns.iLinhasExistentes Then objGridIns.iLinhasExistentes = objGridIns.iLinhasExistentes + 1
        
    End If

    lErro = Grid_Abandona_Celula(objGridIns)
    If lErro <> SUCESSO Then Error 123053
    
    Saida_Celula_ProdutoIns = SUCESSO

    Exit Function
    
Erro_Saida_Celula_ProdutoIns:

    Saida_Celula_ProdutoIns = gErr
    
    Select Case gErr
    
        Case 123053
            Call Grid_Trata_Erro_Saida_Celula(objGridIns)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162674)
            
    End Select

    Exit Function
        
End Function

Private Function Saida_Celula_DescricaoIns(objGridIns As AdmGrid) As Long
'Faz a critica da celula Descricao do grid

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescricaoIns

    Set objGridIns.objControle = DescricaoIns
    
    'Verifica se o Descricao está vazio
    If Len(Trim(DescricaoIns.Text)) <> 0 Then
    
        'Realiza a soma de uma linha no grid
        If GridInsumo.Row - GridInsumo.FixedRows = objGridIns.iLinhasExistentes Then objGridIns.iLinhasExistentes = objGridIns.iLinhasExistentes + 1
        
    End If

    lErro = Grid_Abandona_Celula(objGridIns)
    If lErro <> SUCESSO Then Error 123054
    
    Saida_Celula_DescricaoIns = SUCESSO

    Exit Function
    
Erro_Saida_Celula_DescricaoIns:

    Saida_Celula_DescricaoIns = gErr
    
    Select Case gErr
    
        Case 123054
            Call Grid_Trata_Erro_Saida_Celula(objGridIns)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162675)
            
    End Select

    Exit Function
        
End Function

Private Function Saida_Celula_ParticipacaoIns(objGridIns As AdmGrid) As Long
'Faz a critica da celula Percentual Participacao do grid

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ParticipacaoIns

    Set objGridIns.objControle = PercentualParticipacaoIns
    
    'Verifica se o Percentual Participacao está vazio
    If Len(Trim(PercentualParticipacaoIns.Text)) <> 0 Then
    
        'Realiza a soma de uma linha no grid
        If GridInsumo.Row - GridInsumo.FixedRows = objGridIns.iLinhasExistentes Then objGridIns.iLinhasExistentes = objGridIns.iLinhasExistentes + 1
        
    End If

    lErro = Grid_Abandona_Celula(objGridIns)
    If lErro <> SUCESSO Then Error 123055
    
    Saida_Celula_ParticipacaoIns = SUCESSO

    Exit Function
    
Erro_Saida_Celula_ParticipacaoIns:

    Saida_Celula_ParticipacaoIns = gErr
    
    Select Case gErr
    
        Case 123055
            Call Grid_Trata_Erro_Saida_Celula(objGridIns)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162676)
            
    End Select

    Exit Function
        
End Function

Private Function Saida_Celula_CustoUnitarioIns(objGridIns As AdmGrid) As Long
'Faz a critica da celula Custo Unitario do grid

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_CustoUnitarioIns

    Set objGridIns.objControle = CustoUnitarioIns
    
    'Verifica se o Custo Unitario está vazio
    If Len(Trim(CustoUnitarioIns.ClipText)) <> 0 Then
    
        'Realiza a soma de uma linha no grid
        If GridInsumo.Row - GridInsumo.FixedRows = objGridIns.iLinhasExistentes Then objGridIns.iLinhasExistentes = objGridIns.iLinhasExistentes + 1
        
    End If

    lErro = Grid_Abandona_Celula(objGridIns)
    If lErro <> SUCESSO Then Error 123056
    
    Saida_Celula_CustoUnitarioIns = SUCESSO

    Exit Function
    
Erro_Saida_Celula_CustoUnitarioIns:

    Saida_Celula_CustoUnitarioIns = gErr
    
    Select Case gErr
    
        Case 123056
            Call Grid_Trata_Erro_Saida_Celula(objGridIns)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162677)
            
    End Select

    Exit Function
        
End Function


Private Function Saida_Celula_CustoConsiderarIns(objGridIns As AdmGrid) As Long
'Faz a critica da celula Custo Considerar do grid

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_CustoConsiderarIns

    Set objGridIns.objControle = CustoConsiderarIns
    
    'Verifica se o Custo Considerar está vazio
    If Len(Trim(CustoConsiderarIns.ClipText)) <> 0 Then
    
        'Realiza a soma de uma linha no grid
        If GridInsumo.Row - GridInsumo.FixedRows = objGridIns.iLinhasExistentes Then objGridIns.iLinhasExistentes = objGridIns.iLinhasExistentes + 1
        
    End If

    lErro = Grid_Abandona_Celula(objGridIns)
    If lErro <> SUCESSO Then Error 123057
    
    Saida_Celula_CustoConsiderarIns = SUCESSO

    Exit Function
    
Erro_Saida_Celula_CustoConsiderarIns:

    Saida_Celula_CustoConsiderarIns = gErr
    
    Select Case gErr
    
        Case 123057
            Call Grid_Trata_Erro_Saida_Celula(objGridIns)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162678)
            
    End Select

    Exit Function
        
End Function

Function Saida_Celula_Embalagem(objGridEmb As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Embalagem

    lErro = Grid_Inicializa_Saida_Celula(objGridEmb)
    
    If lErro = SUCESSO Then
        
        'Verifica qual a coluna atual do Grid
        Select Case objGridEmb.objGrid.Col
        
            'Coluna Produto
            Case iGrid_ProdutoEmb_Col
                lErro = Saida_Celula_ProdutoEmb(objGridEmb)
                If lErro <> SUCESSO Then gError 123058
                
            'Coluna Descrição
            Case iGrid_DescricaoEmb_Col
                lErro = Saida_Celula_DescricaoEmb(objGridEmb)
                If lErro <> SUCESSO Then gError 123059
                
            'Coluna Participacao
            Case iGrid_QuantidadeEmb_Col
                lErro = Saida_Celula_QuantidadeEmb(objGridEmb)
                If lErro <> SUCESSO Then gError 123060
                
            'Coluna Custo Unitario
            Case iGrid_CustoUnitarioEmb_Col
                lErro = Saida_Celula_CustoUnitarioEmb(objGridEmb)
                If lErro <> SUCESSO Then gError 123061
                
            'Coluna Custo Considerar
            Case iGrid_CustoConsiderarEmb_Col
                lErro = Saida_Celula_CustoConsiderarEmb(objGridEmb)
                If lErro <> SUCESSO Then gError 123062
                
        End Select
        
        lErro = Grid_Finaliza_Saida_Celula(objGridEmb)
        If lErro <> SUCESSO Then gError 123063

    End If
    
    Saida_Celula_Embalagem = SUCESSO

    Exit Function

Erro_Saida_Celula_Embalagem:

    Saida_Celula_Embalagem = gErr

    Select Case gErr

        Case 123058, 123059, 123060, 123061, 123062, 123063

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162679)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ProdutoEmb(objGridEmb As AdmGrid) As Long
'Faz a critica da celula Produto do grid

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ProdutoEmb

    Set objGridEmb.objControle = ProdutoEmb
    
    'Verifica se o Produto está vazio
    If Len(Trim(ProdutoEmb.ClipText)) <> 0 Then
    
        'Realiza a soma de uma linha no grid
        If GridEmbalagem.Row - GridEmbalagem.FixedRows = objGridEmb.iLinhasExistentes Then objGridEmb.iLinhasExistentes = objGridEmb.iLinhasExistentes + 1
        
    End If

    lErro = Grid_Abandona_Celula(objGridEmb)
    If lErro <> SUCESSO Then Error 123064
    
    Saida_Celula_ProdutoEmb = SUCESSO

    Exit Function
    
Erro_Saida_Celula_ProdutoEmb:

    Saida_Celula_ProdutoEmb = gErr
    
    Select Case gErr
    
        Case 123065
            Call Grid_Trata_Erro_Saida_Celula(objGridEmb)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162680)
            
    End Select

    Exit Function
        
End Function

Private Function Saida_Celula_DescricaoEmb(objGridEmb As AdmGrid) As Long
'Faz a critica da celula Descricao do grid

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescricaoEmb

    Set objGridEmb.objControle = DescricaoEmb
    
    'Verifica se o Descricao está vazio
    If Len(Trim(DescricaoEmb.Text)) <> 0 Then
    
        'Realiza a soma de uma linha no grid
        If GridEmbalagem.Row - GridEmbalagem.FixedRows = objGridEmb.iLinhasExistentes Then objGridEmb.iLinhasExistentes = objGridEmb.iLinhasExistentes + 1
        
    End If

    lErro = Grid_Abandona_Celula(objGridEmb)
    If lErro <> SUCESSO Then Error 123065
    
    Saida_Celula_DescricaoEmb = SUCESSO

    Exit Function
    
Erro_Saida_Celula_DescricaoEmb:

    Saida_Celula_DescricaoEmb = gErr
    
    Select Case gErr
    
        Case 123065
            Call Grid_Trata_Erro_Saida_Celula(objGridEmb)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162681)
            
    End Select

    Exit Function
        
End Function

Private Function Saida_Celula_QuantidadeEmb(objGridEmb As AdmGrid) As Long
'Faz a critica da celula Percentual Quantidade do grid

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_QuantidadeEmb

    Set objGridEmb.objControle = QuantidadeEmb
    
    'Verifica se o Percentual Quantidade está vazio
    If Len(Trim(QuantidadeEmb.ClipText)) <> 0 Then
    
        'Realiza a soma de uma linha no grid
        If GridEmbalagem.Row - GridEmbalagem.FixedRows = objGridEmbalagem.iLinhasExistentes Then objGridEmbalagem.iLinhasExistentes = objGridEmbalagem.iLinhasExistentes + 1
        
    End If

    lErro = Grid_Abandona_Celula(objGridEmb)
    If lErro <> SUCESSO Then Error 123066
    
    Saida_Celula_QuantidadeEmb = SUCESSO

    Exit Function
    
Erro_Saida_Celula_QuantidadeEmb:

    Saida_Celula_QuantidadeEmb = gErr
    
    Select Case gErr
    
        Case 123066
            Call Grid_Trata_Erro_Saida_Celula(objGridEmb)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162682)
            
    End Select

    Exit Function
        
End Function

Private Function Saida_Celula_CustoUnitarioEmb(objGridEmb As AdmGrid) As Long
'Faz a critica da celula Custo Unitario do grid

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_CustoUnitarioEmb

    Set objGridEmb.objControle = CustoUnitarioEmb
    
    'Verifica se o Custo Unitario está vazio
    If Len(Trim(CustoUnitarioEmb.ClipText)) <> 0 Then
    
        'Realiza a soma de uma linha no grid
        If GridEmbalagem.Row - GridEmbalagem.FixedRows = objGridEmb.iLinhasExistentes Then objGridEmb.iLinhasExistentes = objGridEmb.iLinhasExistentes + 1
        
    End If

    lErro = Grid_Abandona_Celula(objGridEmb)
    If lErro <> SUCESSO Then Error 123067
    
    Saida_Celula_CustoUnitarioEmb = SUCESSO

    Exit Function
    
Erro_Saida_Celula_CustoUnitarioEmb:

    Saida_Celula_CustoUnitarioEmb = gErr
    
    Select Case gErr
    
        Case 123067
            Call Grid_Trata_Erro_Saida_Celula(objGridEmb)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162683)
            
    End Select

    Exit Function
        
End Function

Private Function Saida_Celula_CustoConsiderarEmb(objGridEmb As AdmGrid) As Long
'Faz a critica da celula Custo Considerar do grid

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_CustoConsiderarEmb

    Set objGridEmb.objControle = CustoConsiderarEmb
    
    'Verifica se o Custo Considerar está vazio
    If Len(Trim(CustoConsiderarEmb.ClipText)) <> 0 Then
    
        'Realiza a soma de uma linha no grid
        If GridEmbalagem.Row - GridEmbalagem.FixedRows = objGridEmb.iLinhasExistentes Then objGridEmb.iLinhasExistentes = objGridEmb.iLinhasExistentes + 1
        
    End If

    lErro = Grid_Abandona_Celula(objGridEmb)
    If lErro <> SUCESSO Then Error 123068
    
    Saida_Celula_CustoConsiderarEmb = SUCESSO

    Exit Function
    
Erro_Saida_Celula_CustoConsiderarEmb:

    Saida_Celula_CustoConsiderarEmb = gErr
    
    Select Case gErr
    
        Case 123068
            Call Grid_Trata_Erro_Saida_Celula(objGridEmb)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162684)
            
    End Select

    Exit Function
        
End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objKit As New ClassKit

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Kit"

    'Lê os dados da Tela Kit
    lErro = Move_Tela_Memoria(objKit)
    If lErro <> SUCESSO Then gError 123030

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "ProdutoRaiz", objKit.sProdutoRaiz, STRING_PRODUTO, "ProdutoRaiz"
    colCampoValor.Add "Versao", objKit.sVersao, STRING_KIT_VERSAO, "Versao"
    colCampoValor.Add "Data", objKit.dtData, 0, "Data"
    colCampoValor.Add "Observacao", objKit.sObservacao, STRING_KIT_OBSERVACAO, "Observacao"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 123030

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162685)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objKit As New ClassKit

On Error GoTo Erro_Tela_Preenche

    'Preenche o objKit
    objKit.sProdutoRaiz = colCampoValor.Item("ProdutoRaiz").vValor
    objKit.sVersao = colCampoValor.Item("Versao").vValor
    objKit.sObservacao = colCampoValor.Item("Observacao").vValor
    objKit.dtData = colCampoValor.Item("Data").vValor
        
    'Traz dados do Produto_Kit para a Tela
    lErro = Traz_Tela_Produto_MatPrim(objKit)
    If lErro <> SUCESSO And lErro <> 123033 Then gError 123036

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 123036

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162686)

    End Select

    Exit Sub

End Sub

Function Move_Tela_Memoria(objKit As ClassKit) As Long
'Move as informacoes da tela

    'Preenche o objKit com as informacoes da tela
    objKit.sProdutoRaiz = Produto.ClipText
    objKit.sVersao = Versao.Caption
    If Len(Data.Caption) > 0 Then objKit.dtData = CDate(Data.Caption)
    objKit.sObservacao = Observacao.Caption

    Move_Tela_Memoria = SUCESSO

    Exit Function

End Function

Function Trata_Parametros(Optional objKit As ClassKit) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objKit Is Nothing) Then

        lErro = Traz_Tela_Produto_MatPrim(objKit)
        If lErro <> SUCESSO And lErro <> 123033 Then Error 123069

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 123069

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162687)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Traz_Tela_Produto_MatPrim(objKit As ClassKit) As Long
'traz para a tela dados da composicao de um produto
'se a versao nao estiver prenchida traz o padrao

Dim lErro As Long, colMP As New Collection, colEmb As New Collection, iLinha As Integer
Dim sProdutoEnxuto As String, objCustoDirFabrPlanMP As ClassCustoDirFabrPlanMP
Dim objCustoDirFabrPlanEmb As ClassCustoDirFabrPlanEmb, dTotalIns As Double, dTotalEmb As Double
Dim colMPOrd As New Collection, colCamposMP As New Collection, objProduto As New ClassProduto, dTotalPartic As Double
Dim dPercentualPerdaBase As Double, objProdutoKit As New ClassProdutoKit, dParticipacaoTela As Double, dSubtotalIns As Double
Dim objProdutoFilial As New ClassProdutoFilial

On Error GoTo Erro_Traz_Tela_Produto_MatPrim

    lErro = CF("Traz_Produto_MaskEd", objKit.sProdutoRaiz, Produto, Descricao)
    If lErro <> SUCESSO Then Error 123033
    
    objProdutoFilial.iFilialEmpresa = giFilialEmpresa
    objProdutoFilial.sProduto = objKit.sProdutoRaiz
    
    lErro = CF("ProdutoFilial_Le", objProdutoFilial)
    If lErro <> SUCESSO And lErro <> 28261 Then gError 130340
    'Se não encontrou
    If lErro = 28261 Then gError 130341
    
    'se o produto nao é produzido na filial entao pegar apenas o custo dele mesmo como materia prima
    If objProdutoFilial.iProdNaFilial = 0 Then
    
        Versao.Caption = ""
        Data.Caption = ""
        Observacao.Caption = "*** ESTE PRODUTO NÃO É PRODUZIDO NESTA FILIAL ***"
        
    Else
    
        If objKit.sVersao = "" Then
        
            'busca a versao padrao
            lErro = CF("Kit_Le_FormPreco", objKit)
            If lErro <> SUCESSO And lErro <> 106304 Then gError 123031
    
            If lErro = 106304 Then gError 123032
            
        End If
    
        Versao.Caption = objKit.sVersao
        
        objProdutoKit.sProdutoRaiz = objKit.sProdutoRaiz
        objProdutoKit.sVersao = objKit.sVersao
        lErro = CF("ProdutoKit_Le_Raiz", objProdutoKit)
        If lErro <> SUCESSO And lErro <> 34875 Then gError 124283
        
        dPercentualPerdaBase = objProdutoKit.dPercentualPerda
        
        lErro = CF("CustoProd_AjustarPerda", objKit.sProdutoRaiz, objKit.sVersao, giFilialEmpresa, FORMACAO_PRECO_ANALISE_MARGCONTR, dPercentualPerdaBase)
        If lErro <> SUCESSO Then gError 124284
    
        'Preenche a observacao e a descricao
        Observacao.Caption = objKit.sObservacao
        Data.Caption = Format(objKit.dtData, "dd/mm/yy")
    
    End If
    
    'ler do bd a colecao de insumos e embalagens
    lErro = CF("Produto_ObterEmbMP", objKit.sProdutoRaiz, giFilialEmpresa, colMP, colEmb)
    If lErro <> SUCESSO Then gError 106897
    
    'ordena colecao de mps
    For Each objCustoDirFabrPlanMP In colMP
    
        '??? estou usando este campo p/poder usar a rotina de ordenacao de colecoes. O -1 é p/ficar decrescente
        objCustoDirFabrPlanMP.dQtde = -1 * objCustoDirFabrPlanMP.dParticipacao * objCustoDirFabrPlanMP.dCustoUnitario
        
    Next
    
    colCamposMP.Add "dQtde"
    Call Ordena_Colecao(colMP, colMPOrd, colCamposMP)
    
    'preencher grid de insumos
    
    Call Grid_Limpa(objGridInsumo)
    
    iLinha = 0
    For Each objCustoDirFabrPlanMP In colMPOrd
    
        iLinha = iLinha + 1
        
        'Formata o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objCustoDirFabrPlanMP.sProdutoMP, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 106896

        ProdutoIns.PromptInclude = False
        ProdutoIns.Text = sProdutoEnxuto
        ProdutoIns.PromptInclude = True

        'Preenche o Grid
        GridInsumo.TextMatrix(iLinha, iGrid_ProdutoIns_Col) = ProdutoIns.Text
        
        objProduto.sCodigo = objCustoDirFabrPlanMP.sProdutoMP
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 106899
        If lErro <> SUCESSO Then gError 106900
        
        GridInsumo.TextMatrix(iLinha, iGrid_DescricaoIns_Col) = objProduto.sDescricao
        
        dParticipacaoTela = objCustoDirFabrPlanMP.dParticipacao * (1 - dPercentualPerdaBase)
        
        GridInsumo.TextMatrix(iLinha, iGrid_ParticipacaoIns_Col) = Format(dParticipacaoTela, "##0.00%")
        GridInsumo.TextMatrix(iLinha, iGrid_CustoUnitarioIns_Col) = Format(objCustoDirFabrPlanMP.dCustoUnitario, FORMATO_CUSTO)
        GridInsumo.TextMatrix(iLinha, iGrid_CustoConsiderarIns_Col) = Format(dParticipacaoTela * objCustoDirFabrPlanMP.dCustoUnitario, FORMATO_CUSTO)
    
        dTotalPartic = dTotalPartic + dParticipacaoTela
        dSubtotalIns = dSubtotalIns + (dParticipacaoTela * objCustoDirFabrPlanMP.dCustoUnitario)
        dTotalIns = dTotalIns + (objCustoDirFabrPlanMP.dParticipacao * objCustoDirFabrPlanMP.dCustoUnitario)
        
    Next
    
    LabelTotalPartic.Caption = Format(100 * dTotalPartic, "#0.#0\%")
    LabelSubtotalIns.Caption = Format(dSubtotalIns, FORMATO_CUSTO)
    LabelPerda.Caption = Format(100 * (1 - dPercentualPerdaBase), "#0.#0\%")
    LabelTotalIns.Caption = Format(dTotalIns, FORMATO_CUSTO)
    objGridInsumo.iLinhasExistentes = colMPOrd.Count
    
    'preencher grid de embalagens
    
    Call Grid_Limpa(objGridEmbalagem)
    
    iLinha = 0
    For Each objCustoDirFabrPlanEmb In colEmb
    
        iLinha = iLinha + 1
        
        'Formata o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objCustoDirFabrPlanEmb.sProdutoEmb, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 106896

        ProdutoEmb.PromptInclude = False
        ProdutoEmb.Text = sProdutoEnxuto
        ProdutoEmb.PromptInclude = True

        'Preenche o Grid
        GridEmbalagem.TextMatrix(iLinha, iGrid_ProdutoEmb_Col) = ProdutoEmb.Text
        
        objProduto.sCodigo = objCustoDirFabrPlanEmb.sProdutoEmb
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 106901
        If lErro <> SUCESSO Then gError 106902
        
        GridEmbalagem.TextMatrix(iLinha, iGrid_DescricaoEmb_Col) = objProduto.sDescricao
        
        GridEmbalagem.TextMatrix(iLinha, iGrid_QuantidadeEmb_Col) = Format(objCustoDirFabrPlanEmb.dQtde, FORMATO_ESTOQUE)
        GridEmbalagem.TextMatrix(iLinha, iGrid_CustoUnitarioEmb_Col) = Format(objCustoDirFabrPlanEmb.dCustoUnitario, FORMATO_CUSTO)
        GridEmbalagem.TextMatrix(iLinha, iGrid_CustoConsiderarEmb_Col) = Format(objCustoDirFabrPlanEmb.dQtde * objCustoDirFabrPlanEmb.dCustoUnitario, FORMATO_CUSTO)
    
        dTotalEmb = dTotalEmb + (objCustoDirFabrPlanEmb.dQtde * objCustoDirFabrPlanEmb.dCustoUnitario)
        
    Next
    
    LabelTotalEmb.Caption = Format(dTotalEmb, FORMATO_CUSTO)
    objGridEmbalagem.iLinhasExistentes = colEmb.Count
    
    iAlterado = 0

    Traz_Tela_Produto_MatPrim = SUCESSO

    Exit Function

Erro_Traz_Tela_Produto_MatPrim:

    Traz_Tela_Produto_MatPrim = gErr

    Select Case gErr

        Case 123031, 123032, 123033, 124283, 124284, 130340

        Case 130341 'Produto não cadastrado em ProdutoFilial
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FILIAL_INEXISTENTE", gErr, objProdutoFilial.sProduto, giFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162688)

    End Select

    Exit Function

End Function

Private Sub BotaoCustoEmb_Click()

Dim lErro As Long, objCustoEmb As New ClassCustoEmbMP
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoCustoEmb_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridEmbalagem.Row = 0 Then gError 106903
    
    'Se foi selecionada uma linha que está preenchida
    If GridEmbalagem.Row <= objGridEmbalagem.iLinhasExistentes Then
    
        'Verifica se o Produto está preenchido
        lErro = CF("Produto_Formata", GridEmbalagem.TextMatrix(GridEmbalagem.Row, iGrid_ProdutoIns_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 106904
        
        objCustoEmb.sProduto = sProdutoFormatado
        
        Call Chama_Tela("CustoEmbMP", objCustoEmb)
        
    End If
    
    Exit Sub
     
Erro_BotaoCustoEmb_Click:

    Select Case gErr
          
        Case 106903
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 106904
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162689)
     
    End Select
     
    Exit Sub

End Sub

Private Sub BotaoCustoMP_Click()

Dim lErro As Long, objCustoEmb As New ClassCustoEmbMP
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoCustoMP_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridInsumo.Row = 0 Then gError 106903
    
    'Se foi selecionada uma linha que está preenchida
    If GridInsumo.Row <= objGridInsumo.iLinhasExistentes Then
    
        'Verifica se o Produto está preenchido
        lErro = CF("Produto_Formata", GridInsumo.TextMatrix(GridInsumo.Row, iGrid_ProdutoIns_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 106904
        
        objCustoEmb.sProduto = sProdutoFormatado
        
        Call Chama_Tela("CustoEmbMP", objCustoEmb)
        
    End If
    
    Exit Sub
     
Erro_BotaoCustoMP_Click:

    Select Case gErr
          
        Case 106903
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 106904
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162690)
     
    End Select
     
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoDetalhar_Click()
    
Dim lErro As Long
Dim objKit As New ClassKit
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoDetalhar_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridInsumo.Row = 0 Then gError 106903
    
    'Se foi selecionada uma linha que está preenchida
    If GridInsumo.Row <= objGridInsumo.iLinhasExistentes Then
    
        'Verifica se o Produto está preenchido
        lErro = CF("Produto_Formata", GridInsumo.TextMatrix(GridInsumo.Row, iGrid_ProdutoIns_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 106904
        
        objKit.sProdutoRaiz = sProdutoFormatado
    
        'verificar se existe kit p/ele
        lErro = CF("Kit_Le_FormPreco", objKit)
        If lErro <> SUCESSO And lErro <> 106304 Then gError 106905
        If lErro <> SUCESSO Then gError 106906
        
        Call Chama_Tela_Nova_Instancia(Me.Name, objKit)
        
    End If
    
    Exit Sub
     
Erro_BotaoDetalhar_Click:

    Select Case gErr
          
        Case 106903
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 106904, 106905
        
        Case 106906
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_KIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162691)
     
    End Select
     
    Exit Sub

End Sub

Private Sub GridEmbalagem_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridEmbalagem, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridEmbalagem, iAlterado)
    End If

End Sub

Private Sub GridEmbalagem_EnterCell()

     Call Grid_Entrada_Celula(objGridEmbalagem, iAlterado)
     
End Sub

Private Sub GridEmbalagem_GotFocus()

    Call Grid_Recebe_Foco(objGridEmbalagem)

End Sub

Private Sub GridEmbalagem_KeyDown(KeyCode As Integer, Shift As Integer)
 
    Call Grid_Trata_Tecla1(KeyCode, objGridEmbalagem)

    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub
    
End Sub

Private Sub GridEmbalagem_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridEmbalagem, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridEmbalagem, iAlterado)
    End If

End Sub

Private Sub GridEmbalagem_LeaveCell()

    Call Saida_Celula_Embalagem(objGridEmbalagem)

End Sub

Private Sub GridEmbalagem_RowColChange()

    Call Grid_RowColChange(objGridEmbalagem)
    
End Sub

Private Sub GridEmbalagem_Scroll()

    Call Grid_Scroll(objGridEmbalagem)
    
End Sub

Private Sub GridEmbalagem_Validate(Cancel As Boolean)
        
    Call Grid_Libera_Foco(objGridEmbalagem)

End Sub

Private Sub GridInsumo_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridInsumo, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridInsumo, iAlterado)
    End If

End Sub

Private Sub GridInsumo_EnterCell()

    Call Grid_Entrada_Celula(objGridInsumo, iAlterado)
    
End Sub

Private Sub GridInsumo_GotFocus()

    Call Grid_Recebe_Foco(objGridInsumo)

End Sub

Private Sub GridInsumo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridInsumo)

    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub
        
End Sub

Private Sub GridInsumo_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridInsumo, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridInsumo, iAlterado)
    End If

End Sub

Private Sub GridInsumo_LeaveCell()
    
    Call Saida_Celula_Insumo(objGridInsumo)

End Sub

Private Sub GridInsumo_RowColChange()

    Call Grid_RowColChange(objGridInsumo)

End Sub

Private Sub GridInsumo_Scroll()

    Call Grid_Scroll(objGridInsumo)
    
End Sub

Private Sub GridInsumo_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridInsumo)

End Sub

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

       If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

    End If

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProduto As String
Dim objKit As New ClassKit

On Error GoTo Erro_Produto_Validate
   
    If Len(Trim(Produto.ClipText)) = 0 Then
        Descricao.Caption = ""
        Exit Sub
    End If

    sProduto = Produto.Text

    'Critica o formato do Produto e se existe no BD
    lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 Then gError 123041

    'se o produto não estiver cadastrado ==> erro
    If lErro = 25041 Then gError 123042

    'se o produto for gerencial, não pode fazer parte de um kit
    If objProduto.iGerencial = GERENCIAL Then gError 123043

    Descricao.Caption = objProduto.sDescricao

    'Preenche o Produto Raiz com o codigo do produto passado
    objKit.sProdutoRaiz = objProduto.sCodigo
    
    'Preenche o restante da tela
    lErro = Traz_Tela_Produto_MatPrim(objKit)
    If lErro <> SUCESSO Then gError 123046
    
    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 123041, 123046

        Case 123042
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case 123043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162692)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoLbl_Click()
'abrir browse de produtosMatPrim

Dim lErro As Long
Dim objKit As New ClassKit
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_ProdutoLbl_Click

    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 123070

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
    
    objKit.sProdutoRaiz = sProdutoFormatado
    
    'Lista de produtos
    Call Chama_Tela("MatPrimProdutoLista", colSelecao, objKit, objEventoProduto)
    
    Exit Sub

Exit Sub

Erro_ProdutoLbl_Click:

    Select Case gErr

        Case 123070

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162693)

    End Select

    Exit Sub

End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    
    'Incializa a varavel do browse
    Set objEventoProduto = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then Error 123019
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoIns)
    If lErro <> SUCESSO Then gError 123019
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoEmb)
    If lErro <> SUCESSO Then gError 123019
    
    'Inicializa a variavel dos grids
    Set objGridInsumo = New AdmGrid
    Set objGridEmbalagem = New AdmGrid
    
    'Inicializa os grids
    lErro = Inicializa_Grid_Insumo(objGridInsumo)
    If lErro <> SUCESSO Then gError 123020
    
    lErro = Inicializa_Grid_Embalagem(objGridEmbalagem)
    If lErro <> SUCESSO Then gError 123021
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case 123019, 123020, 123021
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162694)

    End Select
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objGridInsumo = Nothing
    Set objGridEmbalagem = Nothing
    
    Set objEventoProduto = Nothing
    
    Set colComponentes = Nothing
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
    
End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)
'preenche a tela atraves do browse

Dim objKit As New ClassKit
Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objKit = obj1

    lErro = Mascara_RetornaProdutoEnxuto(objKit.sProdutoRaiz, sProduto)
    If lErro <> SUCESSO Then gError 123044

    'Preenche o produto da tela
    Produto.PromptInclude = False
    Produto.Text = sProduto
    Produto.PromptInclude = True
    
    Call Produto_Validate(bSGECancelDummy)
    
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 123044
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objKit.sProdutoRaiz)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162695)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Sub Form_Activate()
    Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Análise de Custos de Insumos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "MatPrim"
    
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

Public Sub Unload(objme As Object)
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
