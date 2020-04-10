VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpDespacho 
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7605
   KeyPreview      =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   7605
   Begin VB.Frame Frame5 
      Caption         =   "Produtos"
      Height          =   1020
      Left            =   165
      TabIndex        =   36
      Top             =   2055
      Width           =   7335
      Begin MSMask.MaskEdBox ProdutoInicial 
         Height          =   315
         Left            =   660
         TabIndex        =   7
         Top             =   210
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoFinal 
         Height          =   315
         Left            =   660
         TabIndex        =   8
         Top             =   570
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelProdutoAte 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   255
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   40
         Top             =   615
         Width           =   435
      End
      Begin VB.Label LabelProdutoDe 
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
         Height          =   255
         Left            =   285
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   39
         Top             =   240
         Width           =   360
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2175
         TabIndex        =   38
         Top             =   210
         Width           =   4815
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2175
         TabIndex        =   37
         Top             =   570
         Width           =   4845
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vendedores"
      Height          =   675
      Left            =   165
      TabIndex        =   34
      Top             =   3135
      Width           =   7335
      Begin VB.OptionButton OptVendIndir 
         Caption         =   "Vendas Indiretas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2220
         TabIndex        =   10
         Top             =   225
         Width           =   2010
      End
      Begin VB.OptionButton OptVendDir 
         Caption         =   "Vendas Diretas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   465
         TabIndex        =   9
         Top             =   225
         Value           =   -1  'True
         Width           =   2340
      End
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   300
         Left            =   5385
         TabIndex        =   11
         Top             =   285
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelVendedor 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
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
         Left            =   4440
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   35
         Top             =   330
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Layout"
      Height          =   630
      Left            =   4395
      TabIndex        =   33
      Top             =   750
      Width           =   3120
      Begin VB.ComboBox Layout 
         Height          =   315
         ItemData        =   "RelOpDespachoPur.ctx":0000
         Left            =   75
         List            =   "RelOpDespachoPur.ctx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   2985
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Endereço"
      Height          =   600
      Left            =   4410
      TabIndex        =   31
      Top             =   1380
      Width           =   3105
      Begin MSMask.MaskEdBox Cidade 
         Height          =   315
         Left            =   1020
         TabIndex        =   6
         Top             =   225
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         PromptChar      =   " "
      End
      Begin VB.Label LabelCidade 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
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
         Left            =   300
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   32
         Top             =   270
         Width           =   660
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Rotas"
      Height          =   1725
      Left            =   165
      TabIndex        =   30
      Top             =   3840
      Width           =   7335
      Begin VB.CommandButton BotaoDesmarcar 
         Caption         =   "Desmarcar Todas"
         Height          =   555
         Left            =   5715
         Picture         =   "RelOpDespachoPur.ctx":003B
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1035
         Width           =   1530
      End
      Begin VB.CommandButton BotaoMarcar 
         Caption         =   "Marcar Todas"
         Height          =   555
         Left            =   5715
         Picture         =   "RelOpDespachoPur.ctx":121D
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   210
         Width           =   1530
      End
      Begin VB.ListBox ListRegioes 
         Height          =   1410
         Left            =   90
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   225
         Width           =   5595
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5415
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpDespachoPur.ctx":2237
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpDespachoPur.ctx":2391
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpDespachoPur.ctx":251B
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpDespachoPur.ctx":2A4D
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FramePedido 
      Caption         =   "Pedido"
      Height          =   1260
      Left            =   180
      TabIndex        =   26
      Top             =   750
      Width           =   1695
      Begin MSMask.MaskEdBox PedidoInicial 
         Height          =   300
         Left            =   630
         TabIndex        =   1
         Top             =   300
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PedidoFinal 
         Height          =   300
         Left            =   630
         TabIndex        =   2
         Top             =   810
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelPedFinal 
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
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   28
         Top             =   840
         Width           =   360
      End
      Begin VB.Label LabelPedInicial 
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
         Left            =   255
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   27
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.Frame FrameFaturamento 
      Caption         =   "Data de Emissão"
      Height          =   1245
      Left            =   2130
      TabIndex        =   21
      Top             =   750
      Width           =   2055
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   1665
         TabIndex        =   22
         Top             =   315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   300
         Left            =   705
         TabIndex        =   3
         Top             =   315
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   300
         Left            =   1665
         TabIndex        =   23
         Top             =   795
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   720
         TabIndex        =   4
         Top             =   810
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label dIni 
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
         Left            =   315
         TabIndex        =   25
         Top             =   345
         Width           =   315
      End
      Begin VB.Label dFim 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   270
         TabIndex        =   24
         Top             =   840
         Width           =   360
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpDespachoPur.ctx":2BCB
      Left            =   810
      List            =   "RelOpDespachoPur.ctx":2BCD
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
      Width           =   2730
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
      Left            =   3570
      Picture         =   "RelOpDespachoPur.ctx":2BCF
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   90
      Width           =   1815
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
      Left            =   120
      TabIndex        =   29
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpDespacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoOp As AdmEvento
Attribute objEventoOp.VB_VarHelpID = -1
Private WithEvents objEventoCidade As AdmEvento
Attribute objEventoCidade.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio
Dim giPedidoInicial As Integer

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoOp = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoCidade = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Layout.ListIndex = 0
    
    Cidade.MaxLength = STRING_CIDADE
        
    lErro = CarregaList_Regioes
    If lErro <> SUCESSO Then gError 207089
    
    Call Define_Padrao
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 207089
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168187)

    End Select

    Exit Sub

End Sub

Private Function Define_Padrao() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Define_Padrao
    
    OptVendDir.Value = True
   
    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = Err

    Select Case Err
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168913)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iIndice As Integer
Dim iIndiceRel As Integer
Dim sListCount As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 38334
            
    'pega parâmetro Pedido Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NPEDIDOINIC", sParam)
    If lErro <> SUCESSO Then Error 38335
    
    PedidoInicial.Text = sParam
    
    'pega parâmetro Pedido Final e exibe
    lErro = objRelOpcoes.ObterParametro("NPEDIDOFIM", sParam)
    If lErro Then Error 38336
    
    PedidoFinal.Text = sParam
    
     'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then Error 38337

    Call DateParaMasked(DataInicial, CDate(sParam))
    'DataInicial.PromptInclude = False
    'DataInicial.Text = sParam
    'DataInicial.PromptInclude = True

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then Error 38338

    Call DateParaMasked(DataFinal, CDate(sParam))
    'DataFinal.PromptInclude = False
    'DataFinal.Text = sParam
    'DataFinal.PromptInclude = True
       
    'Limpa a Lista
    For iIndice = 0 To ListRegioes.ListCount - 1
        ListRegioes.Selected(iIndice) = False
    Next
    
    'Obtem o numero de Regioes selecionados na Lista
    lErro = objRelOpcoes.ObterParametro("NLISTCOUNT", sListCount)
    If lErro <> SUCESSO Then gError 207091
    'Percorre toda a Lista
    
    For iIndice = 0 To ListRegioes.ListCount - 1
        
        If sListCount = "0" Then
            ListRegioes.Selected(iIndice) = True
        Else
            'Percorre todas as Regieos que foram slecionados
            For iIndiceRel = 1 To StrParaInt(sListCount)
                lErro = objRelOpcoes.ObterParametro("NLIST" & SEPARADOR & iIndiceRel, sParam)
                If lErro <> SUCESSO Then gError 207092
                'Se o cliente não foi excluido
                If sParam = Codigo_Extrai(ListRegioes.List(iIndice)) Then
                    'Marca as Regioes que foram gravados
                    ListRegioes.Selected(iIndice) = True
                End If
            Next
        End If
    Next
       
'    lErro = objRelOpcoes.ObterParametro("NVENDEDOR", sParam)
'    If lErro <> SUCESSO Then gError 207092
'
'    If StrParaInt(sParam) <> 0 Then
'        Vendedor.Text = CInt(sParam)
'        Call Vendedor_Validate(bSGECancelDummy)
'    End If
   
    lErro = objRelOpcoes.ObterParametro("TCIDADE", sParam)
    If lErro <> SUCESSO Then gError 207092
    
    Cidade.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("NLAYOUT", sParam)
    If lErro <> SUCESSO Then gError 207092

    If StrParaInt(sParam) <> 0 Then
    
        Call Combo_Seleciona_ItemData(Layout, StrParaInt(sParam))

    End If
    
    'Pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.ObterParametro("NVENDEDOR", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If StrParaInt(sParam) <> 0 Then
        Vendedor.Text = CInt(sParam)
        Call Vendedor_Validate(bSGECancelDummy)
    End If

    lErro = objRelOpcoes.ObterParametro("NTIPOVEND", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If StrParaInt(sParam) = VENDEDOR_DIRETO Then
        OptVendDir.Value = True
    Else
        OptVendIndir.Value = True
    End If
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 38334 To 38338, 207091 To 207092
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168188)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoOp = Nothing
    Set gobjRelatorio = Nothing
    Set objEventoCidade = Nothing
    Set objEventoVendedor = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29882
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 38332

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 38332
        
        Case 29882
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168189)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String, iTipoVend As Integer) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim iIndice As Integer
Dim iAchou As Integer
Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then Error 37839

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then Error 37840

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then Error 37841

    End If
    
    If OptVendDir.Value Then
        iTipoVend = VENDEDOR_DIRETO
    Else
        iTipoVend = VENDEDOR_INDIRETO
    End If
    
    'Pedido inicial não pode ser maior que o Pedido final
    If Trim(PedidoInicial.Text) <> "" And Trim(PedidoFinal.Text) <> "" Then
    
         If CLng(PedidoInicial.Text) > CLng(PedidoFinal.Text) Then gError 38340
         
    End If
    
    'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then gError 38341
    
    End If
    
    iAchou = 0
       
    'Limpa a Lista
    For iIndice = 0 To ListRegioes.ListCount - 1
        If ListRegioes.Selected(iIndice) = True Then
            iAchou = 1
            Exit For
        End If
        
    Next
       
    If iAchou = 0 Then gError 207095
       
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 38340
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_INICIAL_MAIOR", gErr)
            PedidoInicial.SetFocus
            
        Case 38341
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus
            
        Case 207095
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUMA_ROTA_SELECIONADA", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168190)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

    Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 47110
    
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Layout.ListIndex = 0
    
    Call Limpa_ListRegioes
    
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then Error 47110
    
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47110
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168191)

    End Select

    Exit Sub
    
End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iIndice As Integer
Dim sRegiao As String
Dim sListCount As String
Dim iNRegiao As Integer
Dim lNumIntRel As Long
Dim lNumIntRelVend As Long
Dim colRegiao As New Collection
Dim bTodasRegioes As Boolean
Dim sProd_I As String
Dim sProd_F As String
Dim iTipoVend As Integer

On Error GoTo Erro_PreencherRelOp
    
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F, iTipoVend)
    If lErro <> SUCESSO Then gError 38344
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 38345
                
    lErro = objRelOpcoes.IncluirParametro("NPEDIDOINIC", PedidoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 38346
    
    lErro = objRelOpcoes.IncluirParametro("NPEDIDOFIM", PedidoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 38347
    
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 38348

    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 38349
    
'    iNRegiao = 1
'    'Percorre toda a Lista
'    For iIndice = 0 To ListRegioes.ListCount - 1
'        If ListRegioes.Selected(iIndice) = True Then
'            sRegiao = Codigo_Extrai(ListRegioes.List(iIndice))
'            'Inclui todas as Regioes que foram slecionados
'            lErro = objRelOpcoes.IncluirParametro("NLIST" & SEPARADOR & iNRegiao, sRegiao)
'            If lErro <> AD_BOOL_TRUE Then gError 207093
'            iNRegiao = iNRegiao + 1
'        End If
'    Next
'    sListCount = iNRegiao - 1
'    'Inclui o numero de Clientes selecionados na Lista
'    lErro = objRelOpcoes.IncluirParametro("NLISTCOUNT", sListCount)
'    If lErro <> AD_BOOL_TRUE Then gError 207094
    
    lErro = objRelOpcoes.IncluirParametro("TCIDADE", Cidade.Text)
    If lErro <> AD_BOOL_TRUE Then gError 207094

'    lErro = objRelOpcoes.IncluirParametro("TVENDEDOR", Vendedor.Text)
'    If lErro <> AD_BOOL_TRUE Then gError 207094
'
'    lErro = objRelOpcoes.IncluirParametro("NVENDEDOR", Codigo_Extrai(Vendedor.Text))
'    If lErro <> AD_BOOL_TRUE Then gError 207094
    
    lErro = objRelOpcoes.IncluirParametro("NLAYOUT", Codigo_Extrai(Layout.Text))
    If lErro <> AD_BOOL_TRUE Then gError 207094
    If bExecutando Then
    
        lErro = RetiraNomes_Sel(colRegiao)
        If lErro <> SUCESSO Then gError 207102
    
        lErro = CF("RelDespacho_Insere", colRegiao, lNumIntRel)
        If lErro <> SUCESSO Then gError 207103
        
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 106964
        
    Else
    
        bTodasRegioes = True
        sListCount = "0"
        For iIndice = 0 To ListRegioes.ListCount - 1
            If Not ListRegioes.Selected(iIndice) Then
                bTodasRegioes = False
                Exit For
            End If
        Next
        
        If Not bTodasRegioes Then
            iNRegiao = 1
            'Percorre toda a Lista
            For iIndice = 0 To ListRegioes.ListCount - 1
                If ListRegioes.Selected(iIndice) = True Then
                    sRegiao = Codigo_Extrai(ListRegioes.List(iIndice))
                    'Inclui todas as Regioes que foram slecionados
                    lErro = objRelOpcoes.IncluirParametro("NLIST" & SEPARADOR & iNRegiao, sRegiao)
                    If lErro <> AD_BOOL_TRUE Then gError 207093
                    iNRegiao = iNRegiao + 1
                End If
            Next
            sListCount = iNRegiao - 1
        End If
        
        'Inclui o numero de Clientes selecionados na Lista
        lErro = objRelOpcoes.IncluirParametro("NLISTCOUNT", sListCount)
        If lErro <> AD_BOOL_TRUE Then gError 207094
        
    End If
    
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TVENDEDOR", Vendedor.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("NVENDEDOR", Codigo_Extrai(Vendedor.Text))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("NTIPOVEND", CStr(iTipoVend))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If bExecutando Then
    
        lErro = CF("RelOpFatProdVend_Prepara", iTipoVend, Codigo_Extrai(Vendedor.Text), lNumIntRelVend)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTRELVEND", CStr(lNumIntRelVend))
        If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 38350

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 38344 To 38350

        Case 207093, 207094, 207102, 207103

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168192)

    End Select

    Exit Function

End Function


Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 38351

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 38352

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 47111
    
        DescProdInic.Caption = ""
        DescProdFim.Caption = ""
        
        lErro = Define_Padrao()
        If lErro <> SUCESSO Then Error 47111
        
        ComboOpcoes.Text = ""
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 38351
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 38352, 47111

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168193)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 207104
    
    gobjRelatorio.sNomeTsk = "despapu" & Codigo_Extrai(Layout.Text)

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 207104

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 207106)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 38354

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then Error 38355

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 38356

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 47112
    
    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 38354
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 38355, 38356, 47112

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168195)

    End Select

    Exit Sub

End Sub


Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sProd_I <> "" Then sExpressao = "Prod >= " & Forprint_ConvTexto(sProd_I)

    If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Prod <= " & Forprint_ConvTexto(sProd_F)

    End If
    
    If Trim(PedidoInicial.Text) <> "" Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Ped >= " & Forprint_ConvLong(CLng(PedidoInicial.Text))
        
    End If

    If Trim(PedidoFinal.Text) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Ped <= " & Forprint_ConvLong(CLng(PedidoFinal.Text))

    End If
    
     If Trim(DataInicial.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataInicial.Text))

    End If
    
    If Trim(DataFinal.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataFinal.Text))

    End If
        
        
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168196)

    End Select

    Exit Function

End Function

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub LabelPedInicial_Click()

Dim lErro As Long
Dim objOp As ClassPedidoDeVenda
Dim colSelecao As Collection


On Error GoTo Erro_LabelPedInicial_Click

    giPedidoInicial = 1

    If Len(Trim(PedidoInicial.Text)) <> 0 Then
    
        lErro = Long_Critica(PedidoInicial.Text)
        If lErro <> SUCESSO Then Error 38357
        
        Set objOp = New ClassPedidoDeVenda
        objOp.lCodigo = CLng(PedidoInicial.Text)

    End If

    Call Chama_Tela("PedidoVendaTodosLista", colSelecao, objOp, objEventoOp)
    
    Exit Sub

Erro_LabelPedInicial_Click:

    Select Case Err
    
        Case 38357

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168197)

    End Select

    Exit Sub

End Sub

Private Sub LabelPedFinal_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objOp As ClassPedidoDeVenda

On Error GoTo Erro_LabelPedFinal_Click

    giPedidoInicial = 0

    If Len(Trim(PedidoFinal.Text)) <> 0 Then
    
        lErro = Long_Critica(PedidoFinal.Text)
        If lErro <> SUCESSO Then Error 38358

        Set objOp = New ClassPedidoDeVenda
        objOp.lCodigo = CLng(PedidoFinal.Text)

    End If

    Call Chama_Tela("PedidoVendaTodosLista", colSelecao, objOp, objEventoOp)
   
   Exit Sub

Erro_LabelPedFinal_Click:

    Select Case Err
    
        Case 38358

        Case Else
           lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168198)

    End Select

    Exit Sub

End Sub


Private Sub objEventoOp_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOp As New ClassPedidoDeVenda

On Error GoTo Erro_objEventoOp_evSelecao

    Set objOp = obj1

    If giPedidoInicial = 1 Then

        PedidoInicial.Text = CStr(objOp.lCodigo)
        
    Else

        PedidoFinal.Text = CStr(objOp.lCodigo)

    End If
    
  
    Exit Sub

Erro_objEventoOp_evSelecao:

    Select Case Err

       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168199)

    End Select

    Exit Sub

End Sub


Private Sub PedidoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim ObjPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_PedidoInicial_Validate

    giPedidoInicial = 1
    
    If Len(Trim(PedidoInicial.Text)) > 0 Then
        
        lErro = Long_Critica(PedidoInicial.Text)
        If lErro <> SUCESSO Then Error 38359
    
        ObjPedidoVenda.lCodigo = CLng(PedidoInicial.Text)
        ObjPedidoVenda.iFilialEmpresa = giFilialEmpresa
        
        lErro = CF("PedidoDeVenda_Le_Todos", ObjPedidoVenda)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then Error 38360
            
        'Pedido não está cadastrado
        If lErro <> SUCESSO Then Error 38361
        
    End If
       
    Exit Sub

Erro_PedidoInicial_Validate:

    Cancel = True


    Select Case Err
    
        Case 38359
        
        Case 38360

        Case 38361
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", Err, ObjPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168200)

    End Select

    Exit Sub

End Sub


Private Sub PedidoFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim ObjPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_PedidoFinal_Validate

    giPedidoInicial = 0
    
    If Len(Trim(PedidoFinal.Text)) > 0 Then
    
        lErro = Long_Critica(PedidoFinal.Text)
        If lErro <> SUCESSO Then Error 38362
    
        ObjPedidoVenda.lCodigo = CLng(PedidoFinal.Text)
        ObjPedidoVenda.iFilialEmpresa = giFilialEmpresa
       
        lErro = CF("PedidoDeVenda_Le_Todos", ObjPedidoVenda)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then Error 38363
            
        'Pedido não está cadastrado
        If lErro <> SUCESSO Then Error 38364
        
    End If
       
    Exit Sub

Erro_PedidoFinal_Validate:

    Cancel = True


    Select Case Err
    
        Case 38362
        
        Case 38363
                
        Case 38364
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", Err, ObjPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168201)

    End Select

    Exit Sub

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then Error 38365

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err

        Case 38365

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168202)

    End Select

    Exit Sub

End Sub


Private Sub DataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        sDataInic = DataInicial.Text
        
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then Error 38366

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err

        Case 38366

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168203)

    End Select

    Exit Sub

End Sub


Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 38367

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 38367
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168204)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 38368

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 38368
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168205)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 38369

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 38369
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168206)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 38370

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 38370
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168207)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_DESPACHO
    Set Form_Load_Ocx = Me
    Caption = "Romaneio de Separação"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpDespacho"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is PedidoInicial Then
            Call LabelPedInicial_Click
        ElseIf Me.ActiveControl Is PedidoFinal Then
            Call LabelPedFinal_Click
        ElseIf Me.ActiveControl Is ProdutoInicial Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoFinal Then
            Call LabelProdutoAte_Click
        ElseIf Me.ActiveControl Is Cidade Then
            Call LabelCidade_Click
        ElseIf Me.ActiveControl Is Vendedor Then
            Call LabelVendedor_Click
        End If
    End If

End Sub


Private Sub LabelPedFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPedFinal, Source, X, Y)
End Sub

Private Sub LabelPedFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPedFinal, Button, Shift, X, Y)
End Sub

Private Sub LabelPedInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPedInicial, Source, X, Y)
End Sub

Private Sub LabelPedInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPedInicial, Button, Shift, X, Y)
End Sub

Private Sub dIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dIni, Source, X, Y)
End Sub

Private Sub dIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dIni, Button, Shift, X, Y)
End Sub

Private Sub dFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dFim, Source, X, Y)
End Sub

Private Sub dFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dFim, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Function CarregaList_Regioes() As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_CarregaList_Regioes
    
    'Preenche Combo Regiao
    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê cada codigo e descricao da tabela RegioesVendas
    lErro = CF("Cod_Nomes_Le", "RegioesVendas", "Codigo", "Descricao", STRING_REGIAO_VENDA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 207090

    'preenche a ComboBox Regiao com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        ListRegioes.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        ListRegioes.ItemData(ListRegioes.NewIndex) = objCodigoDescricao.iCodigo
    Next

    CarregaList_Regioes = SUCESSO

    Exit Function

Erro_CarregaList_Regioes:

    CarregaList_Regioes = gErr

    Select Case gErr

        Case 207900

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172566)

    End Select

    Exit Function

End Function

Private Sub BotaoMarcar_Click()
'marcar todos os itens da listbox
Dim iIndice As Integer

    For iIndice = 0 To ListRegioes.ListCount - 1
        ListRegioes.Selected(iIndice) = True
    Next

End Sub

Private Sub BotaoDesmarcar_Click()
'desmarcar todos os itens da listbox
Dim iIndice As Integer

    For iIndice = 0 To ListRegioes.ListCount - 1
        ListRegioes.Selected(iIndice) = False
    Next

End Sub

Sub Limpa_ListRegioes()

Dim iIndice As Integer

    For iIndice = 0 To ListRegioes.ListCount - 1
        ListRegioes.Selected(iIndice) = False
    Next

End Sub

Public Function RetiraNomes_Sel(colRegioes As Collection) As Long
'Retira da combo todos os nomes que não estão selecionados

Dim iIndice As Integer
Dim lCodRegiao As Long

    For iIndice = 0 To ListRegioes.ListCount - 1
        If ListRegioes.Selected(iIndice) = True Then
            lCodRegiao = LCodigo_Extrai(ListRegioes.List(iIndice))
            colRegioes.Add lCodRegiao
        End If
    Next
    
End Function

Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate

    If Len(Trim(Vendedor.Text)) > 0 Then
   
        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(Vendedor, objVendedor, 0)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO2", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169098)

    End Select

End Sub

Private Sub LabelVendedor_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection
    
    'Preenche com o Vendedor da tela
    objVendedor.iCodigo = Codigo_Extrai(Vendedor.Text)
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1
    
    'Preenche campo Vendedor
    Vendedor.Text = CStr(objVendedor.iCodigo)
    Call Vendedor_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub LabelCidade_Click()

Dim objCidade As New ClassCidades
Dim colSelecao As Collection

    objCidade.sDescricao = Cidade.Text

    'Chama a Tela de browse
    Call Chama_Tela("CidadeLista", colSelecao, objCidade, objEventoCidade)

End Sub

Private Sub objEventoCidade_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCidade As ClassCidades

On Error GoTo Erro_objEventoCidade_evSelecao

    Set objCidade = obj1

    Cidade.Text = objCidade.sDescricao

    Me.Show

    Exit Sub

Erro_objEventoCidade_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202974)

    End Select

    Exit Sub

End Sub

Private Sub Cidade_Validate(Cancel As Boolean)

Dim lErro As Long, objCidade As New ClassCidades
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Cidade_Validate

    If Len(Trim(Cidade.Text)) = 0 Then Exit Sub

    objCidade.sDescricao = Cidade.Text
    
    lErro = CF("Cidade_Le_Nome", objCidade)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM

    If lErro <> SUCESSO Then gError 202976

    Exit Sub

Erro_Cidade_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 202976
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CIDADE")
            If vbMsgRes = vbYes Then
                Call Chama_Tela("CidadeCadastro", objCidade)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202977)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82612

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82613

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 82614

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 82612, 82614

        Case 82613
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169038)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82615

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82616

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 82617

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 82615, 82617

        Case 82616
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169039)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoAte_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoAte_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoFinal.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoFinal.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82631

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 82631

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169040)

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
    If Len(ProdutoInicial.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoInicial.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82630

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 82630

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169041)

    End Select

    Exit Sub

End Sub


Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    'giProdInicial = 0

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 37866
    
    If lErro <> SUCESSO Then Error 43241

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 37866

        Case 43241
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168935)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    'giProdInicial = 1

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 37867
    
    If lErro <> SUCESSO Then Error 43242

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 37867

        Case 43242
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168936)

    End Select

    Exit Sub

End Sub
