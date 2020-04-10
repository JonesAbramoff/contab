VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl AtualizacaoPrecoOcx 
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9375
   KeyPreview      =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   9375
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4560
      Index           =   1
      Left            =   210
      TabIndex        =   0
      Top             =   930
      Width           =   8790
      Begin VB.Frame Frame4 
         Caption         =   "Tabelas de Preço a Serem Atualizadas"
         Height          =   2790
         Left            =   360
         TabIndex        =   41
         Top             =   150
         Width           =   7185
         Begin VB.CommandButton BotaoDesselecTodos 
            Caption         =   "Desmarcar Todas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3240
            Picture         =   "AtualizacaoPreco2Ocx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1935
            Width           =   1695
         End
         Begin VB.CommandButton BotaoSelecionaTodos 
            Caption         =   "Marcar Todas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1440
            Picture         =   "AtualizacaoPreco2Ocx.ctx":11E2
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1935
            Width           =   1695
         End
         Begin VB.ListBox ListaTabelas 
            Columns         =   2
            Height          =   1185
            ItemData        =   "AtualizacaoPreco2Ocx.ctx":21FC
            Left            =   840
            List            =   "AtualizacaoPreco2Ocx.ctx":21FE
            Style           =   1  'Checkbox
            TabIndex        =   1
            Top             =   300
            Width           =   4845
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Produtos a Serem Atualizados"
         Height          =   1365
         Left            =   375
         TabIndex        =   36
         Top             =   3135
         Width           =   7200
         Begin MSMask.MaskEdBox ProdutoDe 
            Height          =   315
            Left            =   900
            TabIndex        =   4
            Top             =   315
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoAte 
            Height          =   315
            Left            =   900
            TabIndex        =   5
            Top             =   855
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            PromptChar      =   " "
         End
         Begin VB.Label DescricaoProdutoAte 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2685
            TabIndex        =   40
            Top             =   855
            Width           =   3225
         End
         Begin VB.Label DescricaoProdutoDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2700
            TabIndex        =   39
            Top             =   330
            Width           =   3225
         End
         Begin VB.Label LabelProdutoDe 
            AutoSize        =   -1  'True
            Caption         =   "De:"
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
            Left            =   480
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   38
            Top             =   405
            Width           =   315
         End
         Begin VB.Label LabelProdutoAte 
            AutoSize        =   -1  'True
            Caption         =   "Até:"
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
            Left            =   435
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   37
            Top             =   915
            Width           =   360
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4560
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   900
      Visible         =   0   'False
      Width           =   8790
      Begin VB.Frame FramePercentual 
         BorderStyle     =   0  'None
         Height          =   3270
         Left            =   105
         TabIndex        =   15
         Top             =   1170
         Width           =   8610
         Begin MSMask.MaskEdBox Percentual 
            Height          =   315
            Left            =   6585
            TabIndex        =   9
            Top             =   585
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
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
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataVigencia 
            Height          =   300
            Left            =   3165
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   570
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataVigencia 
            Height          =   300
            Left            =   2010
            TabIndex        =   44
            Top             =   585
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataBase 
            Height          =   300
            Left            =   7770
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   1305
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataBase 
            Height          =   300
            Left            =   6615
            TabIndex        =   47
            Top             =   1305
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelDataBase 
            AutoSize        =   -1  'True
            Caption         =   "Data Base p/ Reajuste:"
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
            Left            =   4455
            TabIndex        =   48
            ToolTipText     =   "Os preços nesta data serão utilizados como base para aplicar o percentual de reajuste"
            Top             =   1380
            Width           =   2025
         End
         Begin VB.Label LabelDataVigencia 
            AutoSize        =   -1  'True
            Caption         =   "Data de Vigência:"
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
            TabIndex        =   45
            Top             =   645
            Width           =   1545
         End
         Begin VB.Label LabelPercentual 
            AutoSize        =   -1  'True
            Caption         =   "Percentual de Reajuste:"
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
            Left            =   4395
            TabIndex        =   16
            Top             =   645
            Width           =   2070
         End
      End
      Begin VB.Frame FramePlanilha 
         BorderStyle     =   0  'None
         Height          =   3270
         Left            =   150
         TabIndex        =   17
         Top             =   1170
         Visible         =   0   'False
         Width           =   8610
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1755
            TabIndex        =   34
            Text            =   "Planilha"
            Top             =   255
            Width           =   4005
         End
         Begin VB.CommandButton SelPlanilha 
            Caption         =   "Selecionar Planilha"
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
            Left            =   6195
            TabIndex        =   33
            Top             =   255
            Width           =   2055
         End
         Begin VB.Frame Frame3 
            Caption         =   "Localização do Codigo do Produto"
            Height          =   1380
            Left            =   45
            TabIndex        =   28
            Top             =   810
            Width           =   2805
            Begin VB.TextBox LinhaProduto 
               Height          =   300
               Left            =   1500
               TabIndex        =   30
               Top             =   435
               Width           =   750
            End
            Begin VB.TextBox ColunaProduto 
               Height          =   300
               Left            =   1515
               TabIndex        =   29
               Top             =   870
               Width           =   750
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Linha Inicial:"
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
               Left            =   345
               TabIndex        =   32
               Top             =   465
               Width           =   1110
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Coluna Inicial:"
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
               Left            =   225
               TabIndex        =   31
               Top             =   900
               Width           =   1230
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Localização do Novo Preço"
            Height          =   1380
            Left            =   3060
            TabIndex        =   23
            Top             =   810
            Width           =   2505
            Begin VB.TextBox ColunaPreco 
               Height          =   300
               Left            =   1515
               TabIndex        =   25
               Top             =   855
               Width           =   750
            End
            Begin VB.TextBox LinhaPreco 
               Height          =   300
               Left            =   1530
               TabIndex        =   24
               Top             =   405
               Width           =   750
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Coluna Inicial:"
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
               Left            =   225
               TabIndex        =   27
               Top             =   900
               Width           =   1230
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Linha Inicial:"
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
               Left            =   345
               TabIndex        =   26
               Top             =   465
               Width           =   1110
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Outro Produto (a cada)"
            Height          =   1380
            Left            =   5790
            TabIndex        =   18
            Top             =   810
            Width           =   2505
            Begin MSMask.MaskEdBox Linhas 
               Height          =   315
               Left            =   1545
               TabIndex        =   19
               Top             =   405
               Width           =   540
               _ExtentX        =   953
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "9999"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox Colunas 
               Height          =   315
               Left            =   1545
               TabIndex        =   20
               Top             =   855
               Width           =   540
               _ExtentX        =   953
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "9999"
               PromptChar      =   " "
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Num. Linhas:"
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
               Left            =   345
               TabIndex        =   22
               Top             =   465
               Width           =   1125
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Num. Colunas:"
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
               Left            =   225
               TabIndex        =   21
               Top             =   900
               Width           =   1245
            End
         End
         Begin MSComDlg.CommonDialog SelecionaPlanilha 
            Left            =   5820
            Top             =   135
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label2 
            Caption         =   "Nome da Planilha:"
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
            Left            =   45
            TabIndex        =   35
            Top             =   285
            Width           =   1590
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Método de Atualização"
         Height          =   750
         Left            =   120
         TabIndex        =   14
         Top             =   210
         Width           =   8280
         Begin VB.OptionButton BotaoPercentual 
            Caption         =   "Por Percentual de Reajuste"
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
            Left            =   165
            TabIndex        =   7
            Top             =   330
            Value           =   -1  'True
            Width           =   2730
         End
         Begin VB.OptionButton BotaoFormacaoPreco 
            Caption         =   "Utilizando Planilha de Formação de Preços"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3390
            TabIndex        =   8
            Top             =   300
            Width           =   4695
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6960
      ScaleHeight     =   495
      ScaleWidth      =   2145
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   2205
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1650
         Picture         =   "AtualizacaoPreco2Ocx.ctx":2200
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1140
         Picture         =   "AtualizacaoPreco2Ocx.ctx":237E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoAtualizar 
         Caption         =   "Atualiza"
         Height          =   375
         Left            =   75
         TabIndex        =   11
         Top             =   75
         Width           =   975
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5130
      Left            =   120
      TabIndex        =   42
      Top             =   480
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   9049
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Método de Atualização"
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
Attribute VB_Name = "AtualizacaoPrecoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iFrameAtual As Integer
Dim iUltimoProduto As Integer
Dim giFocoInicial As Integer

Public WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Public WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

'Constantes para controlar qual o último campo visitado(ProdutoDe ou ProdutoAte)
Private Const PRODUTO_DE = 1
Private Const PRODUTO_ATE = 2

'Constantes públicas dos tabs
Private Const TAB_Selecao = 1
Private Const TAB_Metodo = 2

Private Sub BotaoAtualizar_Click()

Dim lErro As Long, sProdutoDe As String, sProdutoAte As String
Dim iIndice As Integer
Dim iTabela As Integer
Dim dPercentual As Double
Dim objProduto As New ClassProduto
Dim colTabelas As New Collection
Dim sProduto As String
Dim iPreenchidoDe As Integer
Dim iPreenchidoAte As Integer
Dim dtDataVigencia As Date
Dim dtDataBase As Date

On Error GoTo Erro_BotaoAtualizar_Click

    If Len(Trim(DataVigencia.ClipText)) = 0 Then gError 92445
    
    dtDataVigencia = CDate(DataVigencia.Text)

    'Formatar os dois produtos
    'se ambos estiverem preenchidos verificar se o "até" é >= que o "de"
    For iIndice = 0 To ListaTabelas.ListCount - 1
        
        If ListaTabelas.Selected(iIndice) = True Then
            iTabela = ListaTabelas.ItemData(iIndice)
            colTabelas.Add iTabela
        End If
    Next
    
    'Verificar se ao menos uma tabela está marcada p/atualizar
    If colTabelas.Count = 0 Then gError 33726

    'Verifica se o ProdutoDe foi preenchido
    If Len(Trim(ProdutoDe.ClipText)) > 0 Then
        
        'Passa para o formato do BD
        lErro = CF("Produto_Formata", ProdutoDe.Text, sProduto, iPreenchidoDe)
        If lErro <> SUCESSO Then gError 33734

        'Testa se o codigo está preenchido
        If iPreenchidoDe = PRODUTO_PREENCHIDO Then
            sProdutoDe = sProduto
        Else
            sProdutoDe = ""
        End If
        
    End If
    
    'Verifica se o ProdutoAte foi preenchido
    If Len(Trim(ProdutoAte.ClipText)) > 0 Then
        
        'Passa para o formato do BD
        lErro = CF("Produto_Formata", ProdutoAte.Text, sProduto, iPreenchidoAte)
        If lErro <> SUCESSO Then gError 33735

        'Testa se o codigo está preenchido
        If iPreenchidoAte = PRODUTO_PREENCHIDO Then
            sProdutoAte = sProduto
        Else
            sProdutoAte = ""
        End If
        
    End If
    
    If iPreenchidoDe = PRODUTO_PREENCHIDO And iPreenchidoAte = PRODUTO_PREENCHIDO Then
        
        If sProdutoDe > sProdutoAte Then gError 58190
    
    End If

    If BotaoPercentual.Value = True Then

        If Len(Trim(Percentual.Text)) = 0 Then gError 21276
    
        dPercentual = CDbl(Percentual.Text)
    
        If Len(Trim(DataBase.ClipText)) = 0 Then gError 92474
        
        dtDataBase = CDate(DataBase.Text)
    
        lErro = CF("TabelaPrecoItem_AtualizaPrecosPerc", sProdutoDe, sProdutoAte, dPercentual, colTabelas, dtDataVigencia, dtDataBase)
        If lErro <> SUCESSO Then gError 21277
        
    Else
    
        lErro = CF("TabelaPrecoItem_AtualizarPrecosFP", sProdutoDe, sProdutoAte, colTabelas, dtDataVigencia)
        If lErro <> SUCESSO Then gError 92446
        
    End If

    Call Rotina_Aviso(vbOKOnly, "ATUALIZACAO_PRECO_CONCLUIDA")
    
    Exit Sub

Erro_BotaoAtualizar_Click:

    Select Case gErr

        Case 21276
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_NAOPREENCHIDO", gErr)

        Case 33726
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELA_NAO_MARCADA", gErr)
            
        Case 33734, 33735, 21277

        Case 58190
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTODE_MAIOR_PRODUTOATE", gErr)

        Case 92445
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVIGENCIA_NAO_PREENCHIDA", gErr)

        Case 92474
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATABASE_NAO_PREENCHIDA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143160)

    End Select

    Exit Sub

End Sub

Private Sub BotaoDesselecTodos_Click()
'desmarcar todos os itens da listbox
Dim iIndice As Integer

    For iIndice = 0 To ListaTabelas.ListCount - 1
        ListaTabelas.Selected(iIndice) = False
    Next

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoFormacaoPreco_Click()

    Percentual.Enabled = False
    LabelPercentual.Enabled = False
    LabelDataBase.Enabled = False
    DataBase.Enabled = False
    UpDownDataBase.Enabled = False
    
End Sub

Private Sub BotaoLimpar_Click()

    'Funcao generica que limpa campos da tela
    Call Limpa_Tela(Me)
    DescricaoProdutoDe.Caption = ""
    DescricaoProdutoAte.Caption = ""
    Call BotaoSelecionaTodos_Click

    DataVigencia.PromptInclude = False
    DataVigencia.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataVigencia.PromptInclude = True

    DataBase.PromptInclude = False
    DataBase.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataBase.PromptInclude = True


End Sub

Private Sub BotaoPercentual_Click()

    Percentual.Enabled = True
    LabelPercentual.Enabled = True
    LabelDataBase.Enabled = True
    DataBase.Enabled = True
    UpDownDataBase.Enabled = True

End Sub

Private Sub BotaoSelecionaTodos_Click()
'marcar todos os itens da listbox
Dim iIndice As Integer

    For iIndice = 0 To ListaTabelas.ListCount - 1
        ListaTabelas.Selected(iIndice) = True
    Next

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    giFocoInicial = 1
    
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    
    'Chama Carrega_TabelaPreco
    lErro = Carrega_TabelaPreco()
    If lErro <> SUCESSO Then Error 21279

    If ListaTabelas.ListCount = 0 Then Error 21275

    'Inicializa a máscara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoDe)
    If lErro <> SUCESSO Then Error 16953

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoAte)
    If lErro <> SUCESSO Then Error 16954
    
    iUltimoProduto = PRODUTO_DE

    Call BotaoSelecionaTodos_Click
   
    DataVigencia.PromptInclude = False
    DataVigencia.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataVigencia.PromptInclude = True
   
    DataBase.PromptInclude = False
    DataBase.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataBase.PromptInclude = True
   
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 21275
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_INEXISTENTE1", Err)

        Case 21279, 16953, 16954

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143161)

    End Select

    Exit Sub

End Sub

Private Function Carrega_TabelaPreco() As Long
'Carrega a ListBox ListaTabelas

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome

On Error GoTo Erro_Carrega_TabelaPreco

    lErro = CF("Cod_Nomes_Le", "TabelasDePrecoVenda", "Codigo", "Descricao", STRING_TABELAPRECO_DESCRICAO, colCodigoNome)
    If lErro <> SUCESSO Then Error 21281

    'Preenche a ListBox ListaTabelas com os objetos da coleção
    For Each objCodigoNome In colCodigoNome

        ListaTabelas.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        ListaTabelas.ItemData(ListaTabelas.NewIndex) = objCodigoNome.iCodigo

    Next

    Carrega_TabelaPreco = SUCESSO

    Exit Function

Erro_Carrega_TabelaPreco:

    Carrega_TabelaPreco = Err

    Select Case Err

        Case 21281

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143162)

    End Select

    Exit Function

End Function

Private Sub LabelProdutoAte_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoAte_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoAte.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82671

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 82671

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143163)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoDe_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoDe_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoDe.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82672

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 82672

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143164)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82665

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82666
    
    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoAte, DescricaoProdutoAte)
    If lErro <> SUCESSO Then gError 82667

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 82665, 82667

        Case 82666
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143165)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82668

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82669

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoDe, DescricaoProdutoDe)
    If lErro <> SUCESSO Then gError 82670

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 82668, 82670

        Case 82669
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143166)

    End Select

    Exit Sub

End Sub

Private Sub Percentual_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Percentual_Validate

    If Len(Trim(Percentual.Text)) > 0 Then
    
       lErro = Porcentagem_Critica_Nao_Zero(Percentual.Text)
       If lErro <> SUCESSO Then Error 21309
    
    End If
    
    Exit Sub
    
Erro_Percentual_Validate:

    Cancel = True


    Select Case Err

        Case 21309
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143167)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ProdutoAte_Validate

    iUltimoProduto = PRODUTO_ATE
    giFocoInicial = 0
    
    'Se produto não estiver preenchido --> limpa descrição e unidade de medida
    If Len(Trim(ProdutoAte.ClipText)) = 0 Then

        DescricaoProdutoAte.Caption = ""
            
    Else 'Caso esteja preenchido

        lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 58186

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            objProduto.sCodigo = sProdutoFormatado
            
            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 58187
    
            If lErro = 28030 Then Error 58188
            
            If objProduto.iGerencial = GERENCIAL Then Error 58189

            DescricaoProdutoAte.Caption = objProduto.sDescricao
            
        End If
    
    End If
    
    Exit Sub

Erro_ProdutoAte_Validate:
    
    Cancel = True
    
    Select Case Err
        
        Case 58186, 58187
        
        Case 58188
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", ProdutoAte.Text)

            If vbMsgRes = vbYes Then
    
            Call Chama_Tela("Produto", objProduto)

        End If
        
        Case 58189
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", Err, ProdutoAte.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143168)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ProdutoDe_Validate

    iUltimoProduto = PRODUTO_DE
    giFocoInicial = 1
    
    'Se produto não estiver preenchido --> limpa descrição e unidade de medida
    If Len(Trim(ProdutoDe.ClipText)) = 0 Then

        DescricaoProdutoDe.Caption = ""
            
    Else 'Caso esteja preenchido

        lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 58182

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            objProduto.sCodigo = sProdutoFormatado
            
            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 58183
    
            If lErro = 28030 Then Error 58184
            
            If objProduto.iGerencial = GERENCIAL Then Error 58185

            DescricaoProdutoDe.Caption = objProduto.sDescricao
            
        End If
    
    End If
    
    Exit Sub

Erro_ProdutoDe_Validate:
    
    Cancel = True
    
    Select Case Err
        
        Case 58182, 58183
        
        Case 58184
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", ProdutoDe.Text)

            If vbMsgRes = vbYes Then
    
            Call Chama_Tela("Produto", objProduto)

        End If
        
        Case 58185
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", Err, ProdutoDe.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143169)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()
    
    'Se frame selecionado não for o atual
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
        Select Case iFrameAtual
        
            Case TAB_Selecao
                Parent.HelpContextID = IDH_ATUALIZACAO_PRECOS_SELECAO
                    
            Case TAB_Metodo
                Parent.HelpContextID = IDH_ATUALIZACAO_PRECOS_METODO_ATUALIZACAO
        
        End Select
        
    End If

End Sub

Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 143170)

    End Select

    Exit Function

End Function

Private Sub UpDownDataVigencia_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVigencia_DownClick

    'Diminui a DataVigencia em 1 dia
    lErro = Data_Up_Down_Click(DataVigencia, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 92441

    Exit Sub

Erro_UpDownDataVigencia_DownClick:

    Select Case gErr

        Case 92441

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143171)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataVigencia_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVigencia_UpClick

    'Aumenta a DataVigencia em 1 dia
    lErro = Data_Up_Down_Click(DataVigencia, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 92442

    Exit Sub

Erro_UpDownDataVigencia_UpClick:

    Select Case gErr

        Case 92442

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143172)

    End Select

    Exit Sub

End Sub

Private Sub DataVigencia_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataVigencia)

End Sub

Private Sub DataVigencia_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataVigencia_Validate

    'Se a DataVigencia está preenchida
    If Len(DataVigencia.ClipText) > 0 Then

        'Verifica se a DataVigencia é válida
        lErro = Data_Critica(DataVigencia.Text)
        If lErro <> SUCESSO Then gError 92443

        'Verifica se a DataVigencia é Menor que a Data Atual
        If CDate(DataVigencia.Text) < gdtDataAtual Then gError 92444

    End If

    Exit Sub

Erro_DataVigencia_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 92443
        
        Case 92444
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVIGENCIA_MENOR_DATAATUAL", gErr, CDate(DataVigencia.Text), gdtDataAtual)
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143173)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataBase_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataBase_DownClick

    'Diminui a DataBase em 1 dia
    lErro = Data_Up_Down_Click(DataBase, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 92472

    Exit Sub

Erro_UpDownDataBase_DownClick:

    Select Case gErr

        Case 92472

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143174)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataBase_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataBase_UpClick

    'Aumenta a DataBase em 1 dia
    lErro = Data_Up_Down_Click(DataBase, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 92473

    Exit Sub

Erro_UpDownDataBase_UpClick:

    Select Case gErr

        Case 92473

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143175)

    End Select

    Exit Sub

End Sub

Private Sub DataBase_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataBase)

End Sub

Private Sub DataBase_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataBase_Validate

    'Se a DataBase está preenchida
    If Len(DataBase.ClipText) > 0 Then

        'Verifica se a DataBase é válida
        lErro = Data_Critica(DataBase.Text)
        If lErro <> SUCESSO Then gError 92470

    End If

    Exit Sub

Erro_DataBase_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 92470
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143176)

    End Select

    Exit Sub

End Sub


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ATUALIZACAO_PRECOS_SELECAO
    Set Form_Load_Ocx = Me
    Caption = "Atualização de Preço"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "AtualizacaoPreco"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is ProdutoDe Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoAte Then
            Call LabelProdutoAte_Click
        End If
        
    End If

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

Private Sub DescricaoProdutoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescricaoProdutoAte, Source, X, Y)
End Sub

Private Sub DescricaoProdutoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescricaoProdutoAte, Button, Shift, X, Y)
End Sub

Private Sub DescricaoProdutoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescricaoProdutoDe, Source, X, Y)
End Sub

Private Sub DescricaoProdutoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescricaoProdutoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelProdutoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProdutoDe, Source, X, Y)
End Sub

Private Sub LabelProdutoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProdutoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelProdutoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProdutoAte, Source, X, Y)
End Sub

Private Sub LabelProdutoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProdutoAte, Button, Shift, X, Y)
End Sub

Private Sub LabelPercentual_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPercentual, Source, X, Y)
End Sub

Private Sub LabelPercentual_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPercentual, Button, Shift, X, Y)
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

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub


Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
   Height = UserControl.Height
End Property

Public Property Get Width() As Long
   Width = UserControl.Width
End Property






