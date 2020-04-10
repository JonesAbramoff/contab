VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl RelOpPrecoGrupoOcx 
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9360
   KeyPreview      =   -1  'True
   ScaleHeight     =   4965
   ScaleWidth      =   9360
   Begin VB.Frame Frame1 
      Caption         =   "Tabela"
      Height          =   705
      Left            =   105
      TabIndex        =   30
      Top             =   585
      Width           =   7500
      Begin VB.ComboBox TabelaPrecos 
         Height          =   315
         Left            =   4860
         TabIndex        =   34
         Top             =   240
         Width           =   2430
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   2580
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   210
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataRef 
         Height          =   315
         Left            =   1605
         TabIndex        =   32
         Top             =   210
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelTab 
         AutoSize        =   -1  'True
         Caption         =   "Tabela de Preços:"
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
         Left            =   3255
         TabIndex        =   35
         Top             =   285
         Width           =   1575
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Data Referência:"
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
         Left            =   105
         TabIndex        =   33
         Top             =   255
         Width           =   1470
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Produtos"
      Height          =   3465
      Index           =   1
      Left            =   105
      TabIndex        =   8
      Top             =   1335
      Width           =   9150
      Begin VB.CheckBox AnaliticoGrade 
         Caption         =   "Analíticos com Grade"
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
         Left            =   6810
         TabIndex        =   38
         Top             =   210
         Width           =   2265
      End
      Begin VB.CheckBox AnaliticoGrade 
         Caption         =   "Grades e Kits de Venda"
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
         Left            =   4110
         TabIndex        =   37
         Top             =   210
         Value           =   1  'Checked
         Width           =   2490
      End
      Begin VB.CheckBox AnaliticoGrade 
         Caption         =   "Analíticos sem Grade"
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
         Left            =   1530
         TabIndex        =   36
         Top             =   210
         Value           =   1  'Checked
         Width           =   2280
      End
      Begin VB.Frame Frame9 
         Caption         =   "Categorias"
         Height          =   2100
         Index           =   2
         Left            =   75
         TabIndex        =   20
         Top             =   1230
         Width           =   4785
         Begin VB.ComboBox ComboCategoriaProduto 
            Height          =   315
            Left            =   255
            TabIndex        =   22
            Top             =   1620
            Width           =   1590
         End
         Begin VB.ComboBox ComboCategoriaProdutoItem 
            Height          =   315
            Left            =   2400
            TabIndex        =   21
            Top             =   1620
            Width           =   2190
         End
         Begin MSFlexGridLib.MSFlexGrid GridCategoria 
            Height          =   1830
            Left            =   75
            TabIndex        =   23
            Top             =   210
            Width           =   4620
            _ExtentX        =   8149
            _ExtentY        =   3228
            _Version        =   393216
            Rows            =   6
            Cols            =   3
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Busca por parte de texto dentro dos campos"
         Height          =   2100
         Index           =   5
         Left            =   4935
         TabIndex        =   9
         Top             =   1230
         Width           =   4140
         Begin VB.TextBox CodigoLike 
            Height          =   312
            Left            =   1935
            MaxLength       =   20
            TabIndex        =   14
            ToolTipText     =   "a% = Começa com a, %a = Termina com a e %a% = Possui a em qualquer parte"
            Top             =   240
            Width           =   2145
         End
         Begin VB.TextBox DescricaoLike 
            Height          =   312
            Left            =   1935
            MaxLength       =   20
            TabIndex        =   13
            ToolTipText     =   "a% = Começa com a, %a = Termina com a e %a% = Possui a em qualquer parte"
            Top             =   600
            Width           =   2145
         End
         Begin VB.TextBox NomeRedLike 
            Height          =   312
            Left            =   1935
            MaxLength       =   20
            TabIndex        =   12
            ToolTipText     =   "a% = Começa com a, %a = Termina com a e %a% = Possui a em qualquer parte"
            Top             =   975
            Width           =   2145
         End
         Begin VB.TextBox ReferenciaLike 
            Height          =   312
            Left            =   1935
            MaxLength       =   20
            TabIndex        =   11
            ToolTipText     =   "a% = Começa com a, %a = Termina com a e %a% = Possui a em qualquer parte"
            Top             =   1335
            Width           =   2145
         End
         Begin VB.TextBox ModeloLike 
            Height          =   312
            Left            =   1935
            MaxLength       =   20
            TabIndex        =   10
            ToolTipText     =   "a% = Começa com a, %a = Termina com a e %a% = Possui a em qualquer parte"
            Top             =   1710
            Width           =   2145
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Modelo LIKE"
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
            Left            =   735
            TabIndex        =   19
            Top             =   1770
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nome Reduzido LIKE"
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
            Left            =   30
            TabIndex        =   18
            Top             =   1050
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descrição LIKE"
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
            Index           =   3
            Left            =   495
            TabIndex        =   17
            Top             =   675
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código LIKE"
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
            Left            =   765
            TabIndex        =   16
            Top             =   300
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Referência LIKE"
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
            Left            =   405
            TabIndex        =   15
            Top             =   1410
            Width           =   1395
         End
      End
      Begin MSMask.MaskEdBox ProdutoPai 
         Height          =   315
         Left            =   1545
         TabIndex        =   24
         Top             =   900
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TipoProduto 
         Height          =   315
         Left            =   1545
         TabIndex        =   25
         Top             =   510
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label DescTipoProduto 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2205
         TabIndex        =   29
         Top             =   510
         Width           =   6795
      End
      Begin VB.Label LblTipoProduto 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1050
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   28
         Top             =   570
         Width           =   450
      End
      Begin VB.Label DescProdPai 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4140
         TabIndex        =   27
         Top             =   900
         Width           =   4860
      End
      Begin VB.Label LabelProdutoPai 
         Caption         =   "Produto Pai:"
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
         Height          =   255
         Left            =   405
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   26
         Top             =   930
         Width           =   1215
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPrecoGrupoOcx.ctx":0000
      Left            =   1875
      List            =   "RelOpPrecoGrupoOcx.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2916
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7665
      Picture         =   "RelOpPrecoGrupoOcx.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   675
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7125
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPrecoGrupoOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPrecoGrupoOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPrecoGrupoOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPrecoGrupoOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Opção:"
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
      Height          =   255
      Index           =   0
      Left            =   1230
      TabIndex        =   7
      Top             =   165
      Width           =   615
   End
End
Attribute VB_Name = "RelOpPrecoGrupoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Dim iFrameAtual As Integer
Dim iAlterado As Integer

Dim objGridCategoria As AdmGrid
Dim iGrid_Categoria_Col As Integer
Dim iGrid_Valor_Col As Integer

Private WithEvents objEventoProdutoPai As AdmEvento
Attribute objEventoProdutoPai.VB_VarHelpID = -1
Private WithEvents objEventoTipoDeProduto As AdmEvento
Attribute objEventoTipoDeProduto.VB_VarHelpID = -1

Private Sub DataRef_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataRef)

End Sub

Private Sub DataRef_Validate(Cancel As Boolean)

Dim sData As String
Dim lErro As Long

On Error GoTo Erro_DataRef_Validate

    If Len(DataRef.ClipText) > 0 Then

        sData = DataRef.Text
        
        lErro = Data_Critica(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_DataRef_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211419)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataRef, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataRef.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211420)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataRef, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataRef.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211421)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Form_Load
   
    Set objGridCategoria = New AdmGrid
    Set objEventoProdutoPai = New AdmEvento
    Set objEventoTipoDeProduto = New AdmEvento
   
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoPai)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Inicializa_Grid_Categoria(objGridCategoria)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Le as categorias de produto
    lErro = CF("CategoriasProduto_Le_Todas", colCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22542 Then gError ERRO_SEM_MENSAGEM

    'Preenche CategoriaProduto
    For Each objCategoriaProduto In colCategoriaProduto
        ComboCategoriaProduto.AddItem objCategoriaProduto.sCategoria
    Next
    
    lErro = Carrega_TabelaPrecos()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Define_Padrao

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211422)

    End Select

    Exit Sub

End Sub

Private Sub Define_Padrao()
'Preenche a tela com as opções padrão de FilialEmpresa

Dim lErro As Long

On Error GoTo Erro_Define_Padrao
       
    If TabelaPrecos.ListCount <> 0 Then TabelaPrecos.ListIndex = 0
    
    Call DateParaMasked(DataRef, gdtDataAtual)
    
    AnaliticoGrade(0).Value = vbChecked
    AnaliticoGrade(1).Value = vbChecked
    
    If gobjCRFAT.iSeparaItensGradePrecoDif = DESMARCADO Then
        AnaliticoGrade(2).Value = vbUnchecked
        AnaliticoGrade(2).Enabled = False
    Else
        AnaliticoGrade(2).Value = vbChecked
        AnaliticoGrade(2).Enabled = True
    End If
                
    Exit Sub

Erro_Define_Padrao:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211423)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long, iIndice As Integer
Dim sParam As String, iNumCat As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    Call Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
   
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODPAI", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoPai, DescProdPai)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'pega parâmetro Valor Final e exibe
    lErro = objRelOpcoes.ObterParametro("TTABPRECO", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    TabelaPrecos.Text = sParam
   
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAREF", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call DateParaMasked(DataRef, CDate(sParam))
   
    lErro = objRelOpcoes.ObterParametro("NTIPOPROD", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If StrParaInt(sParam) <> 0 Then
        TipoProduto.PromptInclude = False
        TipoProduto.Text = sParam
        TipoProduto.PromptInclude = True
        
        Call TipoProduto_Validate(bSGECancelDummy)
    End If
   
    lErro = objRelOpcoes.ObterParametro("NNUMCAT", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iNumCat = StrParaInt(sParam)
    
    For iIndice = 1 To iNumCat
    
        lErro = objRelOpcoes.ObterParametro("TCATEGORIA" & CStr(iIndice), sParam)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col) = sParam
    
        lErro = objRelOpcoes.ObterParametro("TCATITEM" & CStr(iIndice), sParam)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        GridCategoria.TextMatrix(iIndice, iGrid_Valor_Col) = sParam
    
    Next
    objGridCategoria.iLinhasExistentes = iNumCat
    
    lErro = objRelOpcoes.ObterParametro("TCODIGOLK", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    CodigoLike.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("TDESCLK", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    DescricaoLike.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("TNOMEREDLK", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    NomeRedLike.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("TREFLK", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    ReferenciaLike.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("TMODELOLK", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    ModeloLike.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("NTIPOSEMGRADE", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If StrParaInt(sParam) = MARCADO Then
        AnaliticoGrade(0).Value = vbChecked
    Else
        AnaliticoGrade(0).Value = vbUnchecked
    End If
    
    lErro = objRelOpcoes.ObterParametro("NTIPOGRADE", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If StrParaInt(sParam) = MARCADO Then
        AnaliticoGrade(1).Value = vbChecked
    Else
        AnaliticoGrade(1).Value = vbUnchecked
    End If
    
    lErro = objRelOpcoes.ObterParametro("NTIPOCOMGRADE", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If StrParaInt(sParam) = MARCADO Then
        AnaliticoGrade(2).Value = vbChecked
    Else
        AnaliticoGrade(2).Value = vbUnchecked
    End If
 
    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211424)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objGridCategoria = Nothing
    Set objEventoProdutoPai = Nothing
    Set objEventoTipoDeProduto = Nothing
   
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 211425
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case 211425
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211426)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridCategoria)
    
    ComboCategoriaProdutoItem.Clear
    
    TabelaPrecos.Text = ""

    DescProdPai.Caption = ""
    DescTipoProduto.Caption = ""
    
    DataRef.PromptInclude = False
    DataRef.Text = ""
    DataRef.PromptInclude = True
   
    ComboOpcoes.SetFocus

End Sub

Private Function Formata_E_Critica_Parametros(ByVal objTabelaPrecoGrupo As ClassTabelaPrecoGrupo) As Long

Dim lErro As Long
Dim iPreenchido As Integer
Dim sProduto As String, iIndice As Integer
Dim objProdCat As ClassProdutoCategoria

On Error GoTo Erro_Formata_E_Critica_Parametros
    
    'Verifica se a Conta está preenchida
    If Len(Trim(TabelaPrecos.Text)) = 0 Then gError 211427
          
    'Verifica se a Conta está preenchida
    If Len(Trim(DataRef.ClipText)) = 0 Then gError 211428
    
    objTabelaPrecoGrupo.iTabela = Codigo_Extrai(TabelaPrecos.Text)
    objTabelaPrecoGrupo.dtDataRef = StrParaDate(DataRef.Text)
    objTabelaPrecoGrupo.iTipoDeProduto = StrParaInt(TipoProduto.Text)
    objTabelaPrecoGrupo.iFilialEmpresa = giFilialEmpresa
    
    If AnaliticoGrade(0) = vbChecked Then
        objTabelaPrecoGrupo.iAnaliticoSemGrade = MARCADO
    Else
        objTabelaPrecoGrupo.iAnaliticoSemGrade = DESMARCADO
    End If
    
    If AnaliticoGrade(1) = vbChecked Then
        objTabelaPrecoGrupo.iGradeKitVenda = MARCADO
    Else
        objTabelaPrecoGrupo.iGradeKitVenda = DESMARCADO
    End If
    
    If AnaliticoGrade(2) = vbChecked Then
        objTabelaPrecoGrupo.iAnaliticoComGrade = MARCADO
    Else
        objTabelaPrecoGrupo.iAnaliticoComGrade = DESMARCADO
    End If
    
    objTabelaPrecoGrupo.sCodigoLike = CodigoLike.Text
    objTabelaPrecoGrupo.sDescricaoLike = DescricaoLike.Text
    objTabelaPrecoGrupo.sModeloLike = ModeloLike.Text
    objTabelaPrecoGrupo.sNomeRedLike = NomeRedLike.Text
    objTabelaPrecoGrupo.sReferenciaLike = ReferenciaLike.Text
   
    If Len(Trim(ProdutoPai.ClipText)) <> 0 Then
   
        lErro = CF("Produto_Formata", ProdutoPai.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
   
        objTabelaPrecoGrupo.sProdutoPai = sProduto
        
    End If
    
    For iIndice = 1 To objGridCategoria.iLinhasExistentes
        Set objProdCat = New ClassProdutoCategoria
        
        objProdCat.sCategoria = GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col)
        objProdCat.sItem = GridCategoria.TextMatrix(iIndice, iGrid_Valor_Col)
    
        objTabelaPrecoGrupo.colCategorias.Add objProdCat
    Next
          
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 211427
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)
            
        Case 211428
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211429)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela
    Call Define_Padrao

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutando As Boolean) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long, iIndice As Integer, lNumIntRel As Long
Dim objTabelaPrecoGrupo As New ClassTabelaPrecoGrupo
Dim objProdCat As ClassProdutoCategoria

On Error GoTo Erro_PreencherRelOp
       
    lErro = Formata_E_Critica_Parametros(objTabelaPrecoGrupo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TPRODPAI", objTabelaPrecoGrupo.sProdutoPai)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("NTABPRECO", CStr(objTabelaPrecoGrupo.iTabela))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TTABPRECO", TabelaPrecos.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
        
    lErro = objRelOpcoes.IncluirParametro("DDATAREF", CStr(objTabelaPrecoGrupo.dtDataRef))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("NTIPOPROD", CStr(objTabelaPrecoGrupo.iTipoDeProduto))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TTIPOPROD", DescTipoProduto.Caption)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TCODIGOLK", objTabelaPrecoGrupo.sCodigoLike)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TDESCLK", objTabelaPrecoGrupo.sDescricaoLike)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEREDLK", objTabelaPrecoGrupo.sNomeRedLike)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TREFLK", objTabelaPrecoGrupo.sReferenciaLike)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TMODELOLK", objTabelaPrecoGrupo.sModeloLike)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("NTIPOCOMGRADE", CStr(objTabelaPrecoGrupo.iAnaliticoComGrade))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("NTIPOGRADE", CStr(objTabelaPrecoGrupo.iGradeKitVenda))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("NTIPOSEMGRADE", CStr(objTabelaPrecoGrupo.iAnaliticoSemGrade))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    iIndice = 0
    For Each objProdCat In objTabelaPrecoGrupo.colCategorias
        iIndice = iIndice + 1
        lErro = objRelOpcoes.IncluirParametro("TCATEGORIA" & CStr(iIndice), objProdCat.sCategoria)
        If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
        
        lErro = objRelOpcoes.IncluirParametro("TCATITEM" & CStr(iIndice), objProdCat.sItem)
        If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    Next
    lErro = objRelOpcoes.IncluirParametro("NNUMCAT", CStr(objTabelaPrecoGrupo.colCategorias.Count))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If bExecutando Then
    
        lErro = CF("RelPrecoGrupo_Prepara", lNumIntRel, objTabelaPrecoGrupo)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    End If
            
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211430)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 211431

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela
        Call Define_Padrao

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 211431
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211432)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211433)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 211434

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 211434
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211435)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211436)

    End Select

    Exit Function

End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is TipoProduto Then
            Call LblTipoProduto_Click
        ElseIf Me.ActiveControl Is ProdutoPai Then
            Call LabelProdutoPai_Click
        End If
    
    End If

End Sub

Private Function Carrega_TabelaPrecos() As Long
'Carrega a Combo TabelaPrecos

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Carrega_TabelaPrecos

    'lê códigos e descrições da tabela TabelasDePrecos e devolve na coleção
    lErro = CF("Cod_Nomes_Le", "TabelasDePreco", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'preenche a combo
    For Each objCodigoNome In colCodigoDescricao
        
        If objCodigoNome.iCodigo <> 0 Then
            TabelaPrecos.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            TabelaPrecos.ItemData(TabelaPrecos.NewIndex) = objCodigoNome.iCodigo
        End If
    
    Next

    Carrega_TabelaPrecos = SUCESSO

    Exit Function

Erro_Carrega_TabelaPrecos:

    Carrega_TabelaPrecos = gErr

    Select Case gErr

        'Erro já tratado
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211437)

    End Select

    Exit Function

End Function

Private Sub TabelaPrecos_Validate(Cancel As Boolean)
'Busca a descricao com código digitado na Combo

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_TabelaPrecos_Validate

    'se uma opcao da lista estiver selecionada, OK
    If TabelaPrecos.ListIndex <> -1 Then Exit Sub

    If Len(Trim(TabelaPrecos.Text)) = 0 Then Exit Sub

    lErro = Combo_Seleciona(TabelaPrecos, iCodigo)
    If lErro <> SUCESSO Then gError 211438

    Exit Sub

Erro_TabelaPrecos_Validate:

    Cancel = True

    Select Case gErr

        Case 211438
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_CADASTRADA", gErr, iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211439)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoPai_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoPai_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoPai.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoPai.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoPai, "Gerencial = 1")

    Exit Sub

Erro_LabelProdutoPai_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211440)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoPai_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoPai_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoPai, DescProdPai)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Me.Show

    Exit Sub

Erro_objEventoProdutoPai_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211441)

    End Select

    Exit Sub

End Sub

Private Sub objEventoTipoDeProduto_evSelecao(obj1 As Object)

Dim objTipoProduto As ClassTipoDeProduto
Dim bCancel As Boolean

    Set objTipoProduto = obj1

    'coloca na tela o Tipo de Produto Selecionado e dispara o evento LostFocus
    TipoProduto.PromptInclude = False
    TipoProduto.Text = objTipoProduto.iTipo
    TipoProduto.PromptInclude = True
    Call TipoProduto_Validate(bCancel)

    Me.Show

End Sub

Public Sub LblTipoProduto_Click()

Dim objTipoDeProduto As ClassTipoDeProduto
Dim colSelecao As Collection

    Call Chama_Tela("TipoProdutoLista", colSelecao, objTipoDeProduto, objEventoTipoDeProduto)

End Sub

Public Sub TipoProduto_GotFocus()
    Call MaskEdBox_TrataGotFocus(TipoProduto, iAlterado)
End Sub

Public Sub TipoProduto_Validate(Cancel As Boolean)
'Se mudar o tipo trazer dele os defaults para os campos da tela

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto

On Error GoTo Erro_TipoProduto_Validate

    If Len(Trim(TipoProduto.Text)) > 0 Then

        'Critica o valor
        lErro = Inteiro_Critica(TipoProduto.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        objTipoProduto.iTipo = CInt(TipoProduto.Text)
    
        'Lê o tipo
        lErro = CF("TipoDeProduto_Le", objTipoProduto)
        If lErro <> SUCESSO And lErro <> 22531 Then gError ERRO_SEM_MENSAGEM
        
        'Se não encontrar --> Erro
        If lErro = 22531 Then gError 211442
        
        DescTipoProduto.Caption = objTipoProduto.sDescricao
        
    Else
        DescTipoProduto.Caption = ""
    End If
    
    Exit Sub

Erro_TipoProduto_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 211442
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", gErr, TipoProduto.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211443)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoPai_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoPai_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoPai, DescProdPai)
    If lErro <> SUCESSO And lErro <> 27095 Then gError ERRO_SEM_MENSAGEM
    
    If lErro <> SUCESSO Then gError 211444

    Exit Sub

Erro_ProdutoPai_Validate:

    Cancel = True

    Select Case gErr

        Case 211444
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
       
        Case ERRO_SEM_MENSAGEM
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211445)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_Categoria(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Categoria

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Categoria")
    objGridInt.colColuna.Add ("Item")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ComboCategoriaProduto.Name)
    objGridInt.colCampo.Add (ComboCategoriaProdutoItem.Name)

    'Colunas do Grid
    iGrid_Categoria_Col = 1
    iGrid_Valor_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridCategoria

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 21

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridCategoria.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Categoria = SUCESSO

    Exit Function

End Function

Private Sub GridCategoria_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCategoria, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCategoria, iAlterado)
    End If

End Sub

Private Sub GridCategoria_GotFocus()

    Call Grid_Recebe_Foco(objGridCategoria)

End Sub

Private Sub GridCategoria_EnterCell()

    Call Grid_Entrada_Celula(objGridCategoria, iAlterado)

End Sub

Private Sub GridCategoria_LeaveCell()

    Call Saida_Celula(objGridCategoria)

End Sub

Private Sub GridCategoria_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridCategoria)

End Sub

Private Sub GridCategoria_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCategoria, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCategoria, iAlterado)
    End If

End Sub

Private Sub GridCategoria_LostFocus()

    Call Grid_Libera_Foco(objGridCategoria)

End Sub

Private Sub GridCategoria_RowColChange()

    Call Grid_RowColChange(objGridCategoria)

End Sub

Private Sub GridCategoria_Scroll()

    Call Grid_Scroll(objGridCategoria)

End Sub

Public Sub ComboCategoriaProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComboCategoriaProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

End Sub

Public Sub ComboCategoriaProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)

End Sub

Public Sub ComboCategoriaProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = ComboCategoriaProduto
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ComboCategoriaProdutoItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComboCategoriaProdutoItem_GotFocus()

Dim lErro As Long

On Error GoTo Erro_ComboCategoriaProdutoItem_GotFocus

    'Preenche com os ítens relacionados a Categoria correspondente
    Call Trata_ComboCategoriaProdutoItem

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

    Exit Sub

Erro_ComboCategoriaProdutoItem_GotFocus:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211446)

    End Select

    Exit Sub

End Sub

Public Sub ComboCategoriaProdutoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)

End Sub

Public Sub ComboCategoriaProdutoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = ComboCategoriaProdutoItem
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Trata_ComboCategoriaProdutoItem()

Dim lErro As Long
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim iIndice As Integer
Dim sValor As String

On Error GoTo Erro_Trata_ComboCategoriaProdutoItem

    sValor = ComboCategoriaProdutoItem.Text

    ComboCategoriaProdutoItem.Clear

    ComboCategoriaProdutoItem.Text = sValor

    'Se alguém estiver selecionado
    If Len(GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)) > 0 Then

        'Preencher a Combo de Itens desta Categoria
        objCategoriaProduto.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)

        lErro = Carrega_ComboCategoriaProdutoItem(objCategoriaProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        For iIndice = 0 To ComboCategoriaProdutoItem.ListCount - 1
            If ComboCategoriaProdutoItem.List(iIndice) = GridCategoria.Text Then
                ComboCategoriaProdutoItem.ListIndex = iIndice
                Exit For
            End If
        Next

    End If

    Exit Sub

Erro_Trata_ComboCategoriaProdutoItem:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211447)

    End Select
    
    Exit Sub

End Sub

Private Function Carrega_ComboCategoriaProdutoItem(objCategoriaProduto As ClassCategoriaProduto) As Long
'Carrega o Item da Categoria na Combobox

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem

On Error GoTo Erro_Carrega_ComboCategoriaProdutoItem

    'Lê a tabela CategoriaProdutoItem a partir da Categoria
    lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colItensCategoria)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Insere na combo CategoriaProdutoItem
    For Each objCategoriaProdutoItem In colItensCategoria
        'Insere na combo CategoriaProduto
        ComboCategoriaProdutoItem.AddItem objCategoriaProdutoItem.sItem
    Next

    Carrega_ComboCategoriaProdutoItem = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaProdutoItem:

    Carrega_ComboCategoriaProdutoItem = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211448)

    End Select

    Exit Function

End Function


Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        If objGridInt.objGrid.Name = GridCategoria.Name Then

            'Verifica qual a coluna do Grid
            Select Case GridCategoria.Col

                Case iGrid_Categoria_Col

                    lErro = Saida_Celula_Categoria(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                Case iGrid_Valor_Col

                    lErro = Saida_Celula_Valor(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            End Select
            
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 211449

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 211449
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211450)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Categoria(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Saida_Celula_Categoria

    Set objGridInt.objControle = ComboCategoriaProduto

    If Len(Trim(ComboCategoriaProduto.Text)) > 0 Then

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ComboCategoriaProduto)
        If lErro <> SUCESSO Then
        
            'Preenche o objeto com a Categoria
             objCategoriaProduto.sCategoria = ComboCategoriaProduto.Text

             'Lê Categoria De Produto no BD
             lErro = CF("CategoriaProduto_Le", objCategoriaProduto)
             If lErro <> SUCESSO And lErro <> 22540 Then gError ERRO_SEM_MENSAGEM

             If lErro <> SUCESSO Then gError 211451  'Categoria não está cadastrada

        End If

        'Verifica se já existe a categoria no Grid
        For iIndice = 1 To objGridCategoria.iLinhasExistentes

            If iIndice <> GridCategoria.Row Then If GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col) = ComboCategoriaProduto.Text Then gError 208319

        Next

        If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    Else
        
        GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Valor_Col) = ""
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 211452

    Saida_Celula_Categoria = SUCESSO

    Exit Function

Erro_Saida_Celula_Categoria:

    Saida_Celula_Categoria = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 211451  'Categoria não está cadastrada

            'pergunta se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CATEGORIAPRODUTO")

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a Tela "CategoriaProduto"
                Call Chama_Tela("CategoriaProduto", objCategoriaProduto)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 211452
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_JA_SELECIONADA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211453)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Item do grid que está deixando de ser a corrente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colItens As New Collection

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridCategoria.objControle = ComboCategoriaProdutoItem

    If Len(Trim(ComboCategoriaProdutoItem.Text)) > 0 Then

        'se o campo de categoria estiver vazio ==> erro
        If Len(GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)) = 0 Then gError 211454

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ComboCategoriaProdutoItem)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
            objCategoriaProdutoItem.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)
            objCategoriaProdutoItem.sItem = ComboCategoriaProdutoItem.Text

            'Lê Categoria De Produto no BD
            lErro = CF("CategoriaProduto_Le_Item", objCategoriaProdutoItem)
            If lErro <> SUCESSO And lErro <> 22603 Then gError ERRO_SEM_MENSAGEM

            If lErro <> SUCESSO Then gError 211455 'Item da Categoria não está cadastrado

        End If

        If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 211454
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_CATEGORIA_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        

        Case 211455 'Item da Categoria não está cadastrado
            'Se não for perguntar se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CATEGORIAPRODUTOITEM")

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Preenche o objeto com a Categoria
                objCategoriaProduto.sCategoria = ComboCategoriaProduto.Text

                'Chama a Tela "CategoriaProduto"
                Call Chama_Tela("CategoriaProduto", objCategoriaProduto)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211456)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_PRECOS
    Set Form_Load_Ocx = Me
    Caption = "Lista de Preços"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPrecos"
    
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

Private Sub LabelTab_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTab, Source, X, Y)
End Sub

Private Sub LabelTab_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTab, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

