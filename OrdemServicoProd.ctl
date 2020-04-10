VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl OrdemServicoProd 
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9540
   ScaleHeight     =   5685
   ScaleWidth      =   9540
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "'"
      Height          =   4710
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   9255
      Begin VB.Frame Frame2 
         Caption         =   " Material a ser produzido "
         Height          =   3270
         Left            =   75
         TabIndex        =   12
         Top             =   1410
         Width           =   9180
         Begin VB.ComboBox Situacao 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "OrdemServicoProd.ctx":0000
            Left            =   7230
            List            =   "OrdemServicoProd.ctx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   375
            Width           =   975
         End
         Begin MSMask.MaskEdBox QuantProduzida 
            Height          =   255
            Left            =   6195
            TabIndex        =   20
            Top             =   405
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   450
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
         Begin VB.TextBox UnidadeMed 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   4440
            MaxLength       =   50
            TabIndex        =   19
            Top             =   405
            Width           =   600
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   255
            Left            =   5130
            TabIndex        =   18
            Top             =   405
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   450
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
         Begin VB.TextBox DescricaoItem 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   1710
            MaxLength       =   50
            TabIndex        =   17
            Top             =   420
            Width           =   2600
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
            Height          =   330
            Left            =   1935
            TabIndex        =   15
            Top             =   2850
            Width           =   1680
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
            Height          =   330
            Left            =   180
            TabIndex        =   14
            Top             =   2850
            Width           =   1680
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   285
            Left            =   195
            TabIndex        =   16
            Top             =   405
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   503
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridProd 
            Height          =   2325
            Left            =   120
            TabIndex        =   13
            Top             =   330
            Width           =   8970
            _ExtentX        =   15822
            _ExtentY        =   4101
            _Version        =   393216
            Rows            =   6
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Ordem de Serviço "
         Height          =   1275
         Left            =   75
         TabIndex        =   2
         Top             =   120
         Width           =   9180
         Begin VB.TextBox PrestadorServico 
            Height          =   285
            Left            =   2190
            MaxLength       =   6
            TabIndex        =   9
            Top             =   675
            Width           =   3090
         End
         Begin VB.TextBox Codigo 
            Height          =   285
            Left            =   2190
            MaxLength       =   6
            TabIndex        =   4
            Top             =   285
            Width           =   1350
         End
         Begin MSComCtl2.UpDown UpDownData 
            Height          =   300
            Left            =   8040
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   270
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   300
            Left            =   6975
            TabIndex        =   6
            Top             =   270
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Status:"
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
            Index           =   1
            Left            =   6315
            TabIndex        =   10
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label2 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   5700
            TabIndex        =   5
            Top             =   330
            Width           =   1230
         End
         Begin VB.Label StatusOS 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6975
            TabIndex        =   11
            Top             =   660
            Width           =   1305
         End
         Begin VB.Label LabelPrestador 
            AutoSize        =   -1  'True
            Caption         =   "Prestador de Serviço:"
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
            TabIndex        =   8
            Top             =   720
            Width           =   1860
         End
         Begin VB.Label CodigoOSLabel 
            AutoSize        =   -1  'True
            Caption         =   "Código O.S.:"
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
            Left            =   1050
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   3
            Top             =   330
            Width           =   1095
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "'"
      Height          =   4710
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   825
      Visible         =   0   'False
      Width           =   9255
      Begin VB.Frame Frame4 
         Caption         =   " Material para consumo na produção "
         Height          =   4185
         Left            =   105
         TabIndex        =   22
         Top             =   345
         Width           =   9105
         Begin MSMask.MaskEdBox Data 
            Height          =   255
            Left            =   6345
            TabIndex        =   29
            Top             =   435
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.TextBox UnidadeMedCons 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   4560
            MaxLength       =   50
            TabIndex        =   25
            Top             =   405
            Width           =   600
         End
         Begin VB.TextBox DescricaoItemCons 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   1845
            MaxLength       =   50
            TabIndex        =   24
            Top             =   390
            Width           =   2600
         End
         Begin VB.CommandButton BotaoProdutosCons 
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
            Height          =   330
            Left            =   225
            TabIndex        =   23
            Top             =   3690
            Width           =   1680
         End
         Begin MSMask.MaskEdBox QuantidadeCons 
            Height          =   285
            Left            =   5250
            TabIndex        =   26
            Top             =   405
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   503
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
         Begin MSMask.MaskEdBox ProdutoCons 
            Height          =   285
            Left            =   195
            TabIndex        =   27
            Top             =   405
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   503
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridCons 
            Height          =   3150
            Left            =   210
            TabIndex        =   28
            Top             =   420
            Width           =   8625
            _ExtentX        =   15214
            _ExtentY        =   5556
            _Version        =   393216
            Rows            =   11
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6840
      ScaleHeight     =   495
      ScaleWidth      =   2550
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   135
      Width           =   2610
      Begin VB.CommandButton BotaoBaixar 
         Height          =   360
         Left            =   1080
         Picture         =   "OrdemServicoProd.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Baixar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "OrdemServicoProd.ctx":01C6
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   585
         Picture         =   "OrdemServicoProd.ctx":0320
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1590
         Picture         =   "OrdemServicoProd.ctx":04AA
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2085
         Picture         =   "OrdemServicoProd.ctx":09DC
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoDesfazerBaixa 
      Height          =   360
      Left            =   7860
      Picture         =   "OrdemServicoProd.ctx":0B5A
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Baixar"
      Top             =   255
      Visible         =   0   'False
      Width           =   420
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5145
      Left            =   90
      TabIndex        =   0
      Top             =   450
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   9075
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Material Consumido"
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
Attribute VB_Name = "OrdemServicoProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'??? pendencias
'criar Rotina_Grid_Enable
'colocar no dic o browse de ordens de servico, provavelmente criar views
'colocar no menu consultas de OSProd abertas e todas analogas a de Ordem de Producao, ver outras consultas analogas que sejam interessantes
'acertar numeracao de erro e transferir rotinas
'tratar situacao em Gravar_Registro de maneira analoga a da tela de ordem de producao
'incluir tratamento de gcolItemOS analogo ao de gcolItemOP na tela de OP
'provavelmente incluir na classe iNumItens e iNumItensBaixados
'falta tratar BotaoBaixar_Click, Trata_Parametros
'tratar prestador de servico: tudo. browse, validacao,...

'Property Variables:
Dim m_Caption As String
Event Unload()

'DECLARACAO DE VARIAVEIS GLOBAIS
Dim iAlterado As Integer
Dim iCodigoAlterado As Integer

Dim iFrameAtual  As Integer

Dim objGridProd As AdmGrid
Dim iGrid_Produto_Col As Integer
Dim iGrid_DescricaoItem_Col As Integer
Dim iGrid_UnidadeMed_Col  As Integer
Dim iGrid_Quantidade_Col  As Integer
Dim iGrid_QuantidadeProd_Col  As Integer
Dim iGrid_Situacao_Col  As Integer

Dim objGridCons As AdmGrid
Dim iGrid_ProdutoCons_Col As Integer
Dim iGrid_DescricaoItemCons_Col As Integer
Dim iGrid_UnidadeMedCons_Col  As Integer
Dim iGrid_QuantidadeCons_Col  As Integer
Dim iGrid_Data_Col  As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoEstoque As AdmEvento
Attribute objEventoEstoque.VB_VarHelpID = -1
Private WithEvents objEventoProdutoCons As AdmEvento
Attribute objEventoProdutoCons.VB_VarHelpID = -1

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    Parent.HelpContextID = IDH_TIPOS_BLOQUEIO
    Set Form_Load_Ocx = Me
    Caption = "Ordem de Serviço de Produção"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "OrdemServicoProd"
    
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

Private Sub BotaoBaixar_Click()
    '??? falta implementar
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

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoCodigo = Nothing
    Set objEventoProduto = Nothing
    Set objEventoEstoque = Nothing
    Set objEventoProdutoCons = Nothing

    Set objGridProd = Nothing
    Set objGridCons = Nothing
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_Load()

Dim lErro As Long
    
On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    
    Set objEventoCodigo = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoEstoque = New AdmEvento
    Set objEventoProdutoCons = New AdmEvento
    
    Call CargaCombo_Situacao(Situacao)
    
    'para evitar scroll horizontal
    QuantProduzida.Width = 1.4 * QuantProduzida.Width
    Data.Width = Data.Width + 1.5
    DescricaoItemCons.Width = DescricaoItemCons.Width * 1.3
    
    'Inicializa máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 22963
    
    'Inicializa máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoCons)
    If lErro <> SUCESSO Then gError 22963
    
    Quantidade.Format = FORMATO_ESTOQUE
    QuantidadeCons.Format = FORMATO_ESTOQUE
    
    'Coloca a Data Atual na Tela
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    lErro = Inicializa_GridProd(objGridProd)
    If lErro <> SUCESSO Then gError 86557
    
    lErro = Inicializa_GridCons(objGridCons)
    If lErro <> SUCESSO Then gError 86558
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 86557, 86558
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163952)

    End Select

    iAlterado = 0
    
    Exit Sub
    
End Sub

Function Trata_Parametros(Optional objTipo As ClassTipoDeBloqueio) As Long

    Trata_Parametros = SUCESSO

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Function Inicializa_GridProd(objGridInt As AdmGrid) As Long

Dim iIndice As Integer

    Set objGridInt = New AdmGrid

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Quant. Produzida")
    objGridInt.colColuna.Add ("Situação")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (QuantProduzida.Name)
    objGridInt.colCampo.Add (Situacao.Name)

    'Colunas do Grid
    iGrid_Produto_Col = 1
    iGrid_DescricaoItem_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
    iGrid_QuantidadeProd_Col = 5
    iGrid_Situacao_Col = 6
    
    objGridInt.objGrid = GridProd

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    objGridInt.iLinhasVisiveis = 6
    
    'Largura da primeira coluna
    GridProd.ColWidth(0) = 300

    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGridInt)

    Inicializa_GridProd = SUCESSO
    
    Exit Function

End Function

Private Function Inicializa_GridCons(objGridInt As AdmGrid) As Long

Dim iIndice As Integer

    Set objGridInt = New AdmGrid

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Data")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ProdutoCons.Name)
    objGridInt.colCampo.Add (DescricaoItemCons.Name)
    objGridInt.colCampo.Add (UnidadeMedCons.Name)
    objGridInt.colCampo.Add (QuantidadeCons.Name)
    objGridInt.colCampo.Add (Data.Name)

    'Colunas do Grid
    iGrid_ProdutoCons_Col = 1
    iGrid_DescricaoItemCons_Col = 2
    iGrid_UnidadeMedCons_Col = 3
    iGrid_QuantidadeCons_Col = 4
    iGrid_Data_Col = 5
    
    objGridInt.objGrid = GridCons

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    objGridInt.iLinhasVisiveis = 9
    
    'Largura da primeira coluna
    GridCons.ColWidth(0) = 400

    Call Grid_Inicializa(objGridInt)

    Inicializa_GridCons = SUCESSO
    
    Exit Function

End Function


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

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO
    iCodigoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOrdemServicoProd As New ClassOrdemServicoProd
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Codigo_Validate

    'Se não houve alteração nos dados da tela --> sai
    If (iCodigoAlterado = 0) Then Exit Sub

    'Se o Código não foi preenchido --> Sai
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub

    objOrdemServicoProd.sCodigo = Codigo.Text
    objOrdemServicoProd.iFilialEmpresa = giFilialEmpresa

    'tenta ler a OS desejada
    lErro = CF("OrdemServicoProd_Le", objOrdemServicoProd)
    If lErro <> SUCESSO And lErro <> 86562 Then gError 86563
    
    'Se o código não é de uma OS existente --> Sai
    If lErro <> SUCESSO Then Exit Sub

    'Pergunta se deseja trazer os dados da OS pra tela
    vbMsg = Rotina_Aviso(vbYesNo, "AVISO_PREENCHER_TELA")
    
    'Se não quiser --> Sai
    If vbMsg = vbNo Then gError 86564

    'Traz a OS para a tela
    lErro = Traz_Tela_OrdemServicoProd(objOrdemServicoProd)
    If lErro <> SUCESSO Then gError 86565

    'Fecha o Comando de Setas
    Call ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 86563 To 86565
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163953)

    End Select

    Exit Sub

End Sub

Function Traz_Tela_OrdemServicoProd(objOrdemServicoProd As ClassOrdemServicoProd) As Long
'Preenche a tela com os dados da OS passada

Dim lErro As Long

On Error GoTo Erro_Traz_Tela_OrdemServicoProd

    Call Limpa_Tela_OrdemServicoProd(NAO_FECHAR_SETAS)
    
    lErro = CF("ItensOrdemServicoProd_Le", objOrdemServicoProd)
    If lErro <> SUCESSO Then gError 21963
    
    lErro = CF("ItensOrdemServicoProdCons_Le", objOrdemServicoProd)
    If lErro <> SUCESSO Then gError 21963
    
    Codigo.Text = CStr(objOrdemServicoProd.sCodigo)

'''' COloca o Prestador de serviço na Tela
    
    Call DateParaMasked(DataEmissao, objOrdemServicoProd.dtDataEmissao)
    
    If objOrdemServicoProd.iStatus = STATUS_LANCADO Then
        StatusOS.Caption = "ATIVO"
    Else
        StatusOS.Caption = "BAIXADO"
    End If
    
    'preenche o grid
    lErro = Preenche_GridProd(objOrdemServicoProd.colItens)
    If lErro <> SUCESSO Then gError 21972
    
    lErro = Preenche_GridCons(objOrdemServicoProd.colItensCons)
    If lErro <> SUCESSO Then gError 21972
    
    iAlterado = 0
    iCodigoAlterado = 0

    Traz_Tela_OrdemServicoProd = SUCESSO

    Exit Function

Erro_Traz_Tela_OrdemServicoProd:

    Traz_Tela_OrdemServicoProd = gErr

    Select Case gErr

        Case 21963, 21972, 21966, 62633, 82801

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163954)

    End Select

    Exit Function

End Function

Sub Limpa_Tela_OrdemServicoProd(Optional iFechaSetas As Integer = FECHAR_SETAS)
'Limpa a Tela

    If iFechaSetas = FECHAR_SETAS Then
        'Fecha o comando das setas se estiver aberto
         Call ComandoSeta_Fechar(Me.Name)
    End If
    
    Call Limpa_Tela(Me)

    StatusOS.Caption = ""
    
    Call Grid_Limpa(objGridCons)
    Call Grid_Limpa(objGridProd)

    iAlterado = 0
    iCodigoAlterado = 0

    Exit Sub

End Sub

Function Preenche_GridProd(colItensProd As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer, iIndice1 As Integer
Dim objItemOSProd As ClassItemOSProd
Dim sProdutoMascarado As String
Dim objProduto As New ClassProduto

On Error GoTo Erro_Preenche_GridProd
  
    For Each objItemOSProd In colItensProd
    
        objProduto.sCodigo = objItemOSProd.sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 86574
        If lErro <> SUCESSO Then gError 86575
        
        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'Mascara produto
        lErro = Mascara_MascararProduto(objItemOSProd.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 22927

        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True

        GridProd.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridProd.TextMatrix(iIndice, iGrid_DescricaoItem_Col) = objProduto.sDescricao
        GridProd.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemOSProd.sSiglaUM
        GridProd.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemOSProd.dQuantidade)
        GridProd.TextMatrix(iIndice, iGrid_QuantidadeProd_Col) = Formata_Estoque(objItemOSProd.dQuantidadeProd)
        
        'preenche Situação
        
        For iIndice1 = 0 To Situacao.ListCount - 1
            If Situacao.ItemData(iIndice1) = objItemOSProd.iStatus Then
                Situacao.ListIndex = iIndice1
                Exit For
            End If
        Next
            
        GridProd.TextMatrix(iIndice, iGrid_Situacao_Col) = Situacao.Text
    
    Next

    Preenche_GridProd = SUCESSO
    Exit Function

Erro_Preenche_GridProd:

    Preenche_GridProd = gErr
    
    Select Case gErr
    
        Case 86574, 86576
        
        Case 86575
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163955)
            
    End Select
            
    Exit Function
    
End Function

Function Preenche_GridCons(colItensProdCons As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItemOSProdCons As ClassItemOSProdCons
Dim sProdutoMascarado As String
Dim objProduto As New ClassProduto

On Error GoTo Erro_Preenche_GridCons
  
    For Each objItemOSProdCons In colItensProdCons
    
        objProduto.sCodigo = objItemOSProdCons.sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 86574
        If lErro <> SUCESSO Then gError 86575
        
        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'Mascara produto
        lErro = Mascara_MascararProduto(objItemOSProdCons.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 22927

        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True

        GridCons.TextMatrix(iIndice, iGrid_ProdutoCons_Col) = Produto.Text
        GridCons.TextMatrix(iIndice, iGrid_DescricaoItemCons_Col) = objProduto.sDescricao
        GridCons.TextMatrix(iIndice, iGrid_UnidadeMedCons_Col) = objItemOSProdCons.sSiglaUM
        GridCons.TextMatrix(iIndice, iGrid_QuantidadeCons_Col) = Formata_Estoque(objItemOSProdCons.dQuantidade)
        GridCons.TextMatrix(iIndice, iGrid_Data_Col) = Format(objItemOSProdCons.dtData, "dd/mm/yyyy")
    
    Next

    Preenche_GridCons = SUCESSO
    Exit Function

Erro_Preenche_GridCons:

    Preenche_GridCons = gErr
    
    Select Case gErr
    
        Case 86574, 86576
        
        Case 86575
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163956)
            
    End Select
            
    Exit Function
    
End Function

Private Sub CodigoOSLabel_Click()

Dim objOrdemServicoProd As New ClassOrdemServicoProd
Dim colSelecao As New Collection

    'preenche o objOS com o código da tela , se estiver preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then objOrdemServicoProd.sCodigo = Codigo.Text
    
    'lista as OS's
    Call Chama_Tela("OrdemServicoProdLista", colSelecao, objOrdemServicoProd, objEventoCodigo)

    Exit Sub

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOrdemServicoProd As ClassOrdemServicoProd

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objOrdemServicoProd = obj1

    'traz OP para a tela
    lErro = Traz_Tela_OrdemServicoProd(objOrdemServicoProd)
    If lErro <> SUCESSO Then gError 34675

    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 34675

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163957)
    End Select

    Exit Sub

End Sub

Private Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)

End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
'Critica se a Data da OP está preenchida corretamente
Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then gError 22933

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True

    Select Case gErr

        Case 22933

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163958)

    End Select

    Exit Sub

End Sub


Private Sub GridProd_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGridProd, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridProd, iAlterado)
        End If

End Sub

Private Sub GridProd_GotFocus()
    Call Grid_Recebe_Foco(objGridProd)
End Sub

Private Sub GridProd_EnterCell()

    Call Grid_Entrada_Celula(objGridProd, iAlterado)

End Sub

Private Sub GridProd_LeaveCell()
    Call Saida_Celula(objGridProd)
End Sub

Private Sub GridProd_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer


    Call Grid_Trata_Tecla(KeyAscii, objGridProd, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProd, iAlterado)
    End If

End Sub

Private Sub GridProd_RowColChange()

    Call Grid_RowColChange(objGridProd)

End Sub

Private Sub GridProd_Scroll()

    Call Grid_Scroll(objGridProd)

End Sub

Private Sub GridCons_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGridCons, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridCons, iAlterado)
        End If

End Sub

Private Sub GridCons_GotFocus()
    Call Grid_Recebe_Foco(objGridCons)
End Sub

Private Sub GridCons_EnterCell()

    Call Grid_Entrada_Celula(objGridCons, iAlterado)

End Sub

Private Sub GridCons_LeaveCell()
    Call Saida_Celula(objGridCons)
End Sub

Private Sub GridCons_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer


    Call Grid_Trata_Tecla(KeyAscii, objGridCons, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCons, iAlterado)
    End If

End Sub

Private Sub GridCons_RowColChange()

    Call Grid_RowColChange(objGridCons)

End Sub

Private Sub GridCons_Scroll()

    Call Grid_Scroll(objGridCons)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        If objGridInt.objGrid.Name = GridProd.Name Then
            lErro = Saida_Celula_GridProd(objGridInt)
        Else
            lErro = Saida_Celula_GridCons(objGridInt)
        End If
        If lErro <> SUCESSO Then gError 86577

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 21986
    
    End If
    
    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 86577
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163959)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_GridProd(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridProd
    
    Select Case GridProd.Col

        Case iGrid_Produto_Col

            lErro = Saida_Celula_Produto(objGridInt)
            If lErro <> SUCESSO Then gError 86578
            
        Case iGrid_Quantidade_Col

            lErro = Saida_Celula_Quantidade(objGridInt)
            If lErro <> SUCESSO Then gError 86579

        Case iGrid_Situacao_Col

            lErro = Saida_Celula_Padrao(objGridInt, Situacao)
            If lErro <> SUCESSO Then gError 124245

    End Select

    Saida_Celula_GridProd = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_GridProd:

    Saida_Celula_GridProd = gErr
    
    Select Case gErr
        
        Case 86578, 86579
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163960)
            
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'faz a critica da celula de proddduto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim objProdutoFilial As New ClassProdutoFilial

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 106433
    
    'se o produto foi preenchido
    If Len(Trim(Produto.ClipText)) <> 0 Then
        
        lErro = CF("Produto_Critica_Estoque", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25077 Then gError 21987

        If lErro = 25077 Then gError 21988

        'se produto estiver preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                
            'verifica se este produto já foi usado da OS
            lErro = VerificaUso_Produto(objProduto.sCodigo)
            If lErro <> SUCESSO And lErro <> 41316 Then gError 55326

            If lErro = 41316 Then gError 41317

            If objProduto.iPCP = PRODUTO_PCP_NAOPODE Or objProduto.iCompras <> PRODUTO_PRODUZIVEL Then gError 22946

             'Preenche a linha do grid
            lErro = ProdutoLinha_Preenche(objProduto)
            If lErro <> SUCESSO Then gError 22945
            
            If Len(Trim(GridProd.TextMatrix(GridProd.Row, iGrid_Quantidade_Col))) = 0 Then
                
                objProdutoFilial.sProduto = sProdutoFormatado
                objProdutoFilial.iFilialEmpresa = giFilialEmpresa
                
                'Busca o Lote Economico do Produto/FilialEmpresa
                lErro = CF("ProdutoFilial_Le", objProdutoFilial)
                If lErro <> SUCESSO And lErro <> 28261 Then gError 106402
                
                'preenche com o lote econômico (caso exista)
                If objProdutoFilial.dLoteEconomico > 0 Then GridProd.TextMatrix(GridProd.Row, iGrid_Quantidade_Col) = objProdutoFilial.dLoteEconomico
                
            End If
        
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 21985

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 21987, 22945, 41318, 21985, 55326, 106430, 106433, 106402
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 21988
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
            
                objProduto.sCodigo = Produto.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 22946
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL1", gErr, Produto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 41317
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_DUPLICADO", gErr, Produto.Text, Codigo.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163961)

    End Select

    Exit Function

End Function

Private Function VerificaUso_Produto(sCodigo As String) As Long
'Verifica se existem produtos repetidos na OP

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_VerificaUso_Produto

    If objGridProd.iLinhasExistentes > 0 Then

        For iIndice = 1 To objGridProd.iLinhasExistentes

            If GridProd.Row <> iIndice Then

                lErro = CF("Produto_Formata", GridProd.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 22947

                If sProdutoFormatado = sCodigo Then gError 41316
                
            End If

        Next

    End If

    VerificaUso_Produto = SUCESSO

    Exit Function

Erro_VerificaUso_Produto:

    VerificaUso_Produto = gErr

    Select Case gErr

        Case 22947, 41316

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163962)

    End Select
    
    Exit Function

End Function

Private Function ProdutoLinha_Preenche(objProduto As ClassProduto) As Long

On Error GoTo Erro_ProdutoLinha_Preenche

    'Unidade de Medida
    GridProd.TextMatrix(GridProd.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMEstoque

    'Descricao
    GridProd.TextMatrix(GridProd.Row, iGrid_DescricaoItem_Col) = objProduto.sDescricao

    Situacao.ListIndex = ITEMOP_SITUACAO_NORMAL
    GridProd.TextMatrix(GridProd.Row, iGrid_Situacao_Col) = Situacao.Text

    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridProd.Row - GridProd.FixedRows) = objGridProd.iLinhasExistentes Then
        objGridProd.iLinhasExistentes = objGridProd.iLinhasExistentes + 1
    End If

    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163963)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantTotal As Double
Dim objProdutoFilial As New ClassProdutoFilial
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade
    
    'se a quantidade foi preenchida
    If Len(Quantidade.ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 41319
           
        'Coloca o Produto no Formato do Banco de Dados
        lErro = CF("Produto_Formata", GridProd.TextMatrix(GridProd.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 111429
        
        objProdutoFilial.sProduto = sProdutoFormatado
        objProdutoFilial.iFilialEmpresa = giFilialEmpresa
        
        'Busca o Lote Mínino do Produto/FilialEmpresa
        lErro = CF("ProdutoFilial_Le", objProdutoFilial)
        If lErro <> SUCESSO And lErro <> 28261 Then gError 111430
        
        'preenche Verifica se Lote mínimo esta preenchdo (caso exista)
        If objProdutoFilial.dLoteMinimo > 0 Then
        
            If StrParaDbl(Quantidade.Text) < objProdutoFilial.dLoteMinimo Then gError 111432
        
        End If
                    
        Quantidade.Text = Formata_Estoque(Quantidade.Text)
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 21994

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 21994
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 41319
            Quantidade.SetFocus
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 111429, 111430
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 111432
            Call Rotina_Erro(vbOKOnly, "ERRO_QDTPRODUTO_MENOR_LOTEMININO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163964)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_GridCons(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridCons
    
    Select Case GridCons.Col

        Case iGrid_ProdutoCons_Col

            lErro = Saida_Celula_ProdutoCons(objGridInt)
            If lErro <> SUCESSO Then gError 86578
            
        Case iGrid_QuantidadeCons_Col

            lErro = Saida_Celula_QuantidadeCons(objGridInt)
            If lErro <> SUCESSO Then gError 86579
        
        Case iGrid_Data_Col
            lErro = Saida_Celula_Data(objGridInt)
            If lErro <> SUCESSO Then gError 86579
    
    End Select

    Saida_Celula_GridCons = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_GridCons:

    Saida_Celula_GridCons = gErr
    
    Select Case gErr
        
        Case 86578, 86579
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163965)
            
    End Select

    Exit Function

End Function

Private Function Saida_Celula_ProdutoCons(objGridInt As AdmGrid) As Long
'faz a critica da celula de proddduto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_ProdutoCons

    Set objGridInt.objControle = ProdutoCons

    lErro = CF("Produto_Formata", ProdutoCons.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 106433
    
    'se o produto foi preenchido
    If Len(Trim(ProdutoCons.ClipText)) <> 0 Then
        
        lErro = CF("Produto_Critica_Estoque", ProdutoCons.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25077 Then gError 21987

        If lErro = 25077 Then gError 21988

        'se produto estiver preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                
            'verifica se este produto já foi usado da OS
            lErro = VerificaUso_ProdutoData(ProdutoCons.Text, StrParaDate(GridCons.TextMatrix(GridCons.Row, iGrid_Data_Col)))
            If lErro <> SUCESSO And lErro <> 41316 Then gError 55326

            If lErro = 41316 Then gError 41317

             'Preenche a linha do grid
            lErro = ProdutoConsLinha_Preenche(objProduto)
            If lErro <> SUCESSO Then gError 22945
                    
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 21985

    Saida_Celula_ProdutoCons = SUCESSO

    Exit Function

Erro_Saida_Celula_ProdutoCons:

    Saida_Celula_ProdutoCons = gErr

    Select Case gErr

        Case 21987, 22945, 41318, 21985, 55326, 106430, 106433, 106402
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 21988
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
            
                objProduto.sCodigo = Produto.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 22946
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL1", gErr, Produto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 41317
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_DUPLICADO", gErr, Produto.Text, Codigo.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163966)

    End Select

    Exit Function

End Function
Private Function ProdutoConsLinha_Preenche(objProduto As ClassProduto) As Long

On Error GoTo Erro_ProdutoConsLinha_Preenche

    'Unidade de Medida
    GridCons.TextMatrix(GridCons.Row, iGrid_UnidadeMedCons_Col) = objProduto.sSiglaUMEstoque

    'Descricao
    GridCons.TextMatrix(GridCons.Row, iGrid_DescricaoItemCons_Col) = objProduto.sDescricao

    'default com a data atual
    GridCons.TextMatrix(GridCons.Row, iGrid_Data_Col) = Format(gdtDataAtual, "dd/mm/yyyy")
    
    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridCons.Row - GridCons.FixedRows) = objGridCons.iLinhasExistentes Then
        objGridCons.iLinhasExistentes = objGridCons.iLinhasExistentes + 1
    End If

    ProdutoConsLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoConsLinha_Preenche:

    ProdutoConsLinha_Preenche = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163967)

    End Select

    Exit Function

End Function


Private Function Saida_Celula_QuantidadeCons(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantTotal As Double
Dim objProdutoFilial As New ClassProdutoFilial
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Saida_Celula_QuantidadeCons

    Set objGridInt.objControle = Quantidade
    
    'se a quantidade foi preenchida
    If Len(QuantidadeCons.ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(QuantidadeCons.Text)
        If lErro <> SUCESSO Then gError 41319
           
        'Coloca o Produto no Formato do Banco de Dados
        lErro = CF("Produto_Formata", GridCons.TextMatrix(GridCons.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 111429
        
        QuantidadeCons.Text = Formata_Estoque(QuantidadeCons.Text)
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 21994

    Saida_Celula_QuantidadeCons = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantidadeCons:

    Saida_Celula_QuantidadeCons = gErr

    Select Case gErr

        Case 41319
            QuantidadeCons.SetFocus
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 111429
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163968)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Data(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Data

    Set objGridInt.objControle = Data

    'verifica se a data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError 31504

        lErro = VerificaUso_ProdutoData(GridCons.TextMatrix(GridCons.Row, iGrid_ProdutoCons_Col), CDate(Data.Text))
        If lErro <> SUCESSO Then gError 31504

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 31505

    Saida_Celula_Data = SUCESSO

    Exit Function

Erro_Saida_Celula_Data:

    Saida_Celula_Data = gErr

    Select Case gErr

        Case 31504, 31505
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163969)

    End Select

    Exit Function

End Function

Private Function VerificaUso_ProdutoData(sCodigo As String, dtData As Date) As Long
'Verifica se existem produtos repetidos na OP

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_VerificaUso_ProdutoData

    If objGridProd.iLinhasExistentes > 0 Then

        For iIndice = 1 To objGridProd.iLinhasExistentes

            If GridProd.Row <> iIndice Then

                If GridProd.TextMatrix(iIndice, iGrid_Produto_Col) = sCodigo And dtData = StrParaDate(GridCons.TextMatrix(iIndice, iGrid_Data_Col)) Then gError 86580
                
            End If

        Next

    End If

    VerificaUso_ProdutoData = SUCESSO

    Exit Function

Erro_VerificaUso_ProdutoData:

    VerificaUso_ProdutoData = gErr

    Select Case gErr

        Case 86580

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163970)

    End Select
    
    Exit Function

End Function

Private Sub GridProd_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridProd)

End Sub

Private Sub GridProd_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridProd)

End Sub

Private Sub GridCons_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridCons)

End Sub

Private Sub GridCons_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridCons)

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProd)

End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProd)

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProd.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridProd)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProd)

End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProd)

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridProd.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridProd)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ProdutoCons_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoCons_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCons)

End Sub

Private Sub ProdutoCons_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCons)

End Sub

Private Sub ProdutoCons_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCons.objControle = ProdutoCons
    lErro = Grid_Campo_Libera_Foco(objGridCons)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantidadeCons_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantidadeCons_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCons)

End Sub

Private Sub QuantidadeCons_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCons)

End Sub

Private Sub QuantidadeCons_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridCons.objControle = QuantidadeCons
    lErro = Grid_Campo_Libera_Foco(objGridCons)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Data_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCons)

End Sub

Private Sub Data_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCons)

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridCons.objControle = Data
    lErro = Grid_Campo_Libera_Foco(objGridCons)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 22931

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 22931

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163971)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_UpClick

    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 22932

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 22932

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163972)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProdutos_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim colSelecao As Collection

On Error GoTo Erro_BotaoProdutos_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridProd.Row = 0 Then gError 43718

    'Verifica se o Produto está preenchido
    If Len(Trim(GridProd.TextMatrix(GridProd.Row, iGrid_Produto_Col))) > 0 Then
    
        lErro = CF("Produto_Formata", GridProd.TextMatrix(GridProd.Row, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 55325
        
        If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
        
    End If

    objProduto.sCodigo = sProduto

    'Lista de produtos produzíveis
    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProduto)
        
    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr
    
        Case 43718
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 55325
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163973)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    If GridProd.Row <> 0 Then

        lErro = CF("Produto_Formata", GridProd.TextMatrix(GridProd.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 22935

        'Se o produto não estiver preenchido
        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then

            sProdutoMascarado = String(STRING_PRODUTO, 0)

            'mascara produto escolhido
            lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 22936

            'verifica sua existência na OP
            lErro = VerificaUso_Produto(objProduto.sCodigo)
            If lErro <> SUCESSO And lErro <> 41316 Then gError 55327

            If lErro = 41316 Then gError 22958

            
            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 22937

            If lErro = 28030 Then gError 22939

            If objProduto.iPCP = PRODUTO_PCP_NAOPODE Or objProduto.iCompras <> PRODUTO_PRODUZIVEL Then gError 55276

            Produto.PromptInclude = False
            Produto.Text = sProdutoMascarado
            Produto.PromptInclude = True

            If Not (Me.ActiveControl Is Produto) Then

                'preenche produto
                GridProd.TextMatrix(GridProd.Row, iGrid_Produto_Col) = sProdutoMascarado
    
                'Preenche a Linha do Grid
                lErro = ProdutoLinha_Preenche(objProduto)
                If lErro <> SUCESSO Then gError 22938
    
            End If

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 22935, 22937, 22938, 41305, 55327

        Case 22936
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)
            
        Case 22939
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 22958
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_DUPLICADO", gErr, sProdutoMascarado, Codigo.Text)
            
        Case 55276
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL1", gErr, sProdutoMascarado)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163974)

    End Select

    Exit Sub

End Sub
Private Sub BotaoEstoque_Click()
'Informa se produto é estocado em algum almoxarifado

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoEstoque_Click

    If GridProd.Row = 0 Then gError 43719

    sCodProduto = GridProd.TextMatrix(GridProd.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 21930

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        colSelecao.Add sProdutoFormatado
        'chama a tela de lista de estoque do produto corrente
        Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoEstoque)
    Else
        Error 43739
    End If

    Exit Sub

Erro_BotaoEstoque_Click:

    Select Case gErr

        Case 21930
        
        Case 43719
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 43739
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163975)

    End Select

    Exit Sub

End Sub

Private Sub objEventoEstoque_evselecao(obj1 As Object)

   Me.Show

End Sub

Private Sub BotaoProdutosCons_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim colSelecao As Collection

On Error GoTo Erro_BotaoProdutosCons_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridCons.Row = 0 Then gError 43718

    'Verifica se o Produto está preenchido
    If Len(Trim(GridCons.TextMatrix(GridCons.Row, iGrid_Produto_Col))) > 0 Then
    
        lErro = CF("Produto_Formata", GridCons.TextMatrix(GridCons.Row, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 55325
        
        If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
        
    End If

    objProduto.sCodigo = sProduto

    'Lista de produtos produzíveis
    Call Chama_Tela("ProdutoEstoquePCPLista", colSelecao, objProduto, objEventoProdutoCons)
        
    Exit Sub

Erro_BotaoProdutosCons_Click:

    Select Case gErr
    
        Case 43718
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 55325
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163976)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoProdutoCons_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String

On Error GoTo Erro_objEventoProdutoCons_evSelecao

    Set objProduto = obj1

    If GridCons.Row <> 0 Then

        lErro = CF("Produto_Formata", GridCons.TextMatrix(GridCons.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 22935

        'Se o produto não estiver preenchido
        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then

            sProdutoMascarado = String(STRING_PRODUTO, 0)

            'mascara produto escolhido
            lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 22936

            ProdutoCons.PromptInclude = False
            ProdutoCons.Text = sProdutoMascarado
            ProdutoCons.PromptInclude = True
            
            'verifica sua existência na OP
            lErro = VerificaUso_ProdutoData(Produto.Text, StrParaDate(GridCons.TextMatrix(GridCons.Row, iGrid_Data_Col)))
            If lErro <> SUCESSO And lErro <> 41316 Then gError 55327

            If lErro = 41316 Then gError 22958
            
            If Not (Me.ActiveControl Is ProdutoCons) Then

                'preenche produto
                GridCons.TextMatrix(GridCons.Row, iGrid_ProdutoCons_Col) = ProdutoCons.Text
    
                'Preenche a Linha do Grid
                lErro = ProdutoConsLinha_Preenche(objProduto)
                If lErro <> SUCESSO Then gError 22938
    
            End If

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProdutoCons_evSelecao:

    Select Case gErr

        Case 22935, 22937, 22938, 41305, 55327

        Case 22936
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)
            
        Case 22939
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 22958
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_DUPLICADO", gErr, sProdutoMascarado, Codigo.Text)
            
        Case 55276
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL1", gErr, sProdutoMascarado)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163977)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'testa se houva alguma alteração
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 21952

    'limpa a tela
    Call Limpa_Tela_OrdemServicoProd
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 21952

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163978)

    End Select

    Exit Sub

End Sub


Private Sub BotaoGravar_Click()
'implementa gravação de uma nova ou atualizacao de uma OP

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Rotina de gravação da OP
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 21750

    'limpa a tela
    Call Limpa_Tela_OrdemServicoProd
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 21750, 21951

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163979)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objOrdemServico As New ClassOrdemServicoProd

On Error GoTo Erro_Gravar_Registro

    If Len(Trim(Codigo)) = 0 Then gError 11111
    
    If Len(Trim(PrestadorServico.Text)) = 0 Then gError 11112
    
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 11113
    
    If objGridProd.iLinhasExistentes = 0 Then gError 11114
    
    objOrdemServico.sCodigo = Codigo.Text
    objOrdemServico.lCodPrestador = PrestadorServico.Tag
    objOrdemServico.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objOrdemServico.iFilialEmpresa = giFilialEmpresa
    
    lErro = Move_GridProd_Memoria(objOrdemServico)
    If lErro <> SUCESSO Then gError 11115
    
    lErro = Move_GridCons_Memoria(objOrdemServico)
    If lErro <> SUCESSO Then gError 11116

    lErro = OrdemServicoProd_Grava(objOrdemServico)
    If lErro <> SUCESSO Then gError 11117

    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 11111
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INFORMADO1", gErr)
        
        Case 11112
            Call Rotina_Erro(vbOKOnly, "ERRO_PRESTADOR_NAO_INFORMADO", gErr)

        Case 11113
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_INFORMADA", gErr)

        Case 11114
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_GRID", gErr)
        
        Case 11115 To 11117

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163980)

    End Select

    Exit Function
    
End Function

Function Move_GridProd_Memoria(objOrdemServico As ClassOrdemServicoProd) As Long

Dim lErro As Long
Dim iIndice As Integer, iCount As Integer
Dim objItemOSProd As ClassItemOSProd
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sProduto As String, sSituacao As String

On Error GoTo Erro_Move_GridProd_Memoria

    For iIndice = 1 To objGridProd.iLinhasExistentes
                
        Set objItemOSProd = New ClassItemOSProd
        
        sProduto = GridProd.TextMatrix(iIndice, iGrid_Produto_Col)

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 31563

        objItemOSProd.sProduto = sProdutoFormatado
        objItemOSProd.dQuantidade = StrParaDbl(GridProd.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objItemOSProd.iFilialEmpresa = objOrdemServico.iFilialEmpresa
        objItemOSProd.sCodigo = objOrdemServico.sCodigo
        objItemOSProd.iItem = iIndice
        
        'Seleciona a situação
        If Len(Trim(GridProd.TextMatrix(iIndice, iGrid_Situacao_Col))) > 0 Then
            sSituacao = GridProd.TextMatrix(iIndice, iGrid_Situacao_Col)
            For iCount = 0 To Situacao.ListCount - 1
                If Situacao.List(iCount) = sSituacao Then
                    objItemOSProd.iStatus = Situacao.ItemData(iCount)
                    Exit For
                End If
            Next
        End If

        objItemOSProd.sSiglaUM = GridProd.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
        
        objOrdemServico.colItens.Add objItemOSProd
        
    Next

    Move_GridProd_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_GridProd_Memoria:

    Move_GridProd_Memoria = gErr
    
    Select Case gErr
    
        Case 31563
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163981)

    End Select

    Exit Function
    
End Function

Function Move_GridCons_Memoria(objOrdemServico As ClassOrdemServicoProd) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItemOSProdCons As ClassItemOSProdCons
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sProduto As String

On Error GoTo Erro_Move_GridCons_Memoria

    For iIndice = 1 To objGridCons.iLinhasExistentes
                
        Set objItemOSProdCons = New ClassItemOSProdCons
        
        sProduto = GridCons.TextMatrix(iIndice, iGrid_ProdutoCons_Col)

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 31563

        objItemOSProdCons.sProduto = sProdutoFormatado
        objItemOSProdCons.dQuantidade = StrParaDbl(GridCons.TextMatrix(iIndice, iGrid_QuantidadeCons_Col))
        objItemOSProdCons.iFilialEmpresa = objOrdemServico.iFilialEmpresa
        objItemOSProdCons.sCodigo = objOrdemServico.sCodigo
        objItemOSProdCons.iItem = iIndice
        objItemOSProdCons.sSiglaUM = GridCons.TextMatrix(iIndice, iGrid_UnidadeMedCons_Col)
        objItemOSProdCons.dtData = StrParaDate(GridCons.TextMatrix(iIndice, iGrid_Data_Col))
        
        objOrdemServico.colItensCons.Add objItemOSProdCons
        
    Next

    Move_GridCons_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_GridCons_Memoria:

    Move_GridCons_Memoria = gErr
    
    Select Case gErr
    
        Case 31563
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163982)

    End Select

    Exit Function
    
End Function

Function OrdemServicoProd_Grava(objOrdemServico As ClassOrdemServicoProd) As Long

Dim lTransacao As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lErro As Long, iStatus As Integer

On Error GoTo Erro_OrdemServicoProd_Grava

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 11111
    Next

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 11112
    
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Status FROM OrdemServicoProd WHERE Codigo = ? AND FilialEmpresa = ? ", 0, iStatus, objOrdemServico.sCodigo, objOrdemServico.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 11113

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 11114

    If lErro = AD_SQL_SEM_DADOS Then
    
        lErro = Comando_Executar(alComando(0), "INSERT INTO OrdemServicoProd (Codigo, FilialEmpresa, DataEmissao, PrestadorServico, Status) VALUES (?,?,?,?,?) ", objOrdemServico.sCodigo, objOrdemServico.iFilialEmpresa, objOrdemServico.dtDataEmissao, objOrdemServico.lCodPrestador, objOrdemServico.iStatus)
        If lErro <> AD_SQL_SUCESSO Then gError 11115
        
        lErro = ItensOSProd_Grava(objOrdemServico)
        If lErro <> SUCESSO Then gError 11116
        
        lErro = ItensOSProdCons_Grava(objOrdemServico)
        If lErro <> SUCESSO Then gError 11116
        
    Else
    
    End If

    OrdemServicoProd_Grava = SUCESSO
     
    Exit Function
    
Erro_OrdemServicoProd_Grava:

    OrdemServicoProd_Grava = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163983)
     
    End Select
     
    Exit Function

End Function

Function ItensOSProd_Grava(objOrdemServico As ClassOrdemServicoProd) As Long

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim objItemOSProd As ClassItemOSProd

On Error GoTo Erro_ItensOSProd_Grava

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 11110

    For iIndice = 1 To objOrdemServico.colItens.Count
        
        Set objItemOSProd = objOrdemServico.colItens(iIndice)
        
        lNumIntDoc = 0
        
        'Obtem o próximo número interno de RastreamentoLote disponivel
        lErro = CF("Config_ObterNumInt", "MatConfig", "NUM_PROX_INT_ITEMOSPROD", lNumIntDoc)
        If lErro <> SUCESSO Then gError 11111
        
        objItemOSProd.lNumIntDoc = lNumIntDoc
        
        With objItemOSProd
            
            lErro = Comando_Executar(lComando, "INSERT INTO ItensOrdemServicoProd (NumIntDoc, FilialEmpresa, Codigo, Item, Produto, Quantidade, QuantidadeProd, SiglaUM, Status) VALUES (?,?,?,?,?,?,?,?,?)", .lNumIntDoc, .iFilialEmpresa, .sCodigo, .iItem, .sProduto, .dQuantidade, .dQuantidadeProd, .sSiglaUM, .iStatus)
            If lErro <> SUCESSO Then gError 11112
        
        End With
        
    Next
        
    Call Comando_Fechar(lComando)
    
    ItensOSProd_Grava = SUCESSO
    
    Exit Function

Erro_ItensOSProd_Grava:

    ItensOSProd_Grava = gErr
    
    Select Case gErr
        
        Case 11110
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 11111
        
        Case 11112
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSOSPROD", gErr, objOrdemServico.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163984)

    End Select

    Exit Function

End Function

Function ItensOSProdCons_Grava(objOrdemServico As ClassOrdemServicoProd) As Long

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim objItemOSProdCons As ClassItemOSProdCons

On Error GoTo Erro_ItensOSProdCons_Grava

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 11110

    For iIndice = 1 To objOrdemServico.colItens.Count
        
        Set objItemOSProdCons = objOrdemServico.colItens(iIndice)
        
        lNumIntDoc = 0
        
        'Obtem o próximo número interno de RastreamentoLote disponivel
        lErro = CF("Config_ObterNumInt", "MatConfig", "NUM_PROX_INT_ITEMOSPRODCONS", lNumIntDoc)
        If lErro <> SUCESSO Then gError 11111
        
        objItemOSProdCons.lNumIntDoc = lNumIntDoc
        
        With objItemOSProdCons
            
            lErro = Comando_Executar(lComando, "INSERT INTO ItensOrdemServicoProdCons (NumIntDoc, FilialEmpresa, Codigo, Item, Produto, Quantidade, SiglaUM, Data) VALUES (?,?,?,?,?,?,?,?,?)", .lNumIntDoc, .iFilialEmpresa, .sCodigo, .iItem, .sProduto, .dQuantidade, .sSiglaUM, .dtData)
            If lErro <> SUCESSO Then gError 11112
        
        End With
        
    Next
        
    Call Comando_Fechar(lComando)
    
    ItensOSProdCons_Grava = SUCESSO
    
    Exit Function

Erro_ItensOSProdCons_Grava:

    ItensOSProdCons_Grava = gErr
    
    Select Case gErr
        
        Case 11110
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 11111
        
        Case 11112
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSOSPRODCONS", gErr, objOrdemServico.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163985)

    End Select

    Exit Function

End Function

Public Sub CargaCombo_Situacao(objSituacao As Object)
'Carga dos itens da combo Situação

    objSituacao.AddItem STRING_NORMAL
    objSituacao.ItemData(objSituacao.NewIndex) = ITEMOP_SITUACAO_NORMAL
    objSituacao.AddItem STRING_BAIXADA
    objSituacao.ItemData(objSituacao.NewIndex) = ITEMOP_SITUACAO_BAIXADA

End Sub

Private Sub Situacao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Situacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProd)

End Sub

Private Sub Situacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProd)

End Sub

Private Sub Situacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProd.objControle = Situacao
    lErro = Grid_Campo_Libera_Foco(objGridProd)
    If lErro <> SUCESSO Then Cancel = True

End Sub

