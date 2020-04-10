VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl RelOpOPHar 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleMode       =   0  'User
   ScaleWidth      =   9510
   Begin VB.Frame FrameEmissao 
      Caption         =   "Emissão"
      Height          =   765
      Left            =   135
      TabIndex        =   24
      Top             =   840
      Width           =   4200
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   300
         Left            =   480
         TabIndex        =   2
         Top             =   300
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   2685
         TabIndex        =   4
         Top             =   285
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   1500
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   300
         Width           =   240
         _ExtentX        =   397
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   285
         Left            =   3690
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   397
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   -1  'True
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
         Left            =   2280
         TabIndex        =   26
         Top             =   345
         Width           =   360
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
         Left            =   120
         TabIndex        =   25
         Top             =   330
         Width           =   315
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Categorias de Produto"
      Height          =   1860
      Left            =   135
      TabIndex        =   14
      Top             =   4065
      Width           =   9270
      Begin VB.CommandButton BotaoMarTodosCat 
         Caption         =   "Marca Todos"
         Height          =   615
         Left            =   7650
         Picture         =   "RelOpOPHar.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   330
         Width           =   1455
      End
      Begin VB.CommandButton BotaoDesTodosCat 
         Caption         =   "Desmarca Todos"
         Height          =   615
         Left            =   7665
         Picture         =   "RelOpOPHar.ctx":101A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1020
         Width           =   1455
      End
      Begin VB.ListBox ListaCat 
         Columns         =   3
         Height          =   1410
         Left            =   315
         Style           =   1  'Checkbox
         TabIndex        =   19
         Top             =   300
         Width           =   7215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordens de Produção"
      Height          =   2370
      Left            =   135
      TabIndex        =   13
      Top             =   1650
      Width           =   9270
      Begin VB.CommandButton BotaoMarTodosOP 
         Caption         =   "Marca Todos"
         Height          =   615
         Left            =   7695
         Picture         =   "RelOpOPHar.ctx":21FC
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   345
         Width           =   1455
      End
      Begin VB.CommandButton BotaoDesTodosOP 
         Caption         =   "Desmarca Todos"
         Height          =   615
         Left            =   7710
         Picture         =   "RelOpOPHar.ctx":3216
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1035
         Width           =   1455
      End
      Begin VB.ListBox ListaOP 
         Columns         =   3
         Height          =   1860
         Left            =   330
         Style           =   1  'Checkbox
         TabIndex        =   18
         Top             =   345
         Width           =   7200
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filial Empresa"
      Height          =   765
      Left            =   4560
      TabIndex        =   12
      Top             =   840
      Width           =   4830
      Begin VB.ComboBox ComboFilial 
         Height          =   315
         Left            =   1695
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   285
         Width           =   3030
      End
      Begin VB.CheckBox TodasFiliais 
         Caption         =   "Todas"
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
         Left            =   195
         TabIndex        =   15
         Top             =   315
         Width           =   855
      End
      Begin VB.Label FilialLabel 
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
         Left            =   1200
         TabIndex        =   17
         Top             =   345
         Width           =   465
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7230
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpOPHar.ctx":43F8
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpOPHar.ctx":4552
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpOPHar.ctx":46DC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpOPHar.ctx":4C0E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpOPHar.ctx":4D8C
      Left            =   795
      List            =   "RelOpOPHar.ctx":4D8E
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   270
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
      Left            =   4005
      Picture         =   "RelOpOPHar.ctx":4D90
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   105
      Width           =   1575
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
      Left            =   135
      TabIndex        =   11
      Top             =   315
      Width           =   615
   End
End
Attribute VB_Name = "RelOpOPHar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iFilialAnt As Integer
Dim dtDataInicialAnt As Date
Dim dtDataFinalAnt As Date

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Const TAM_COD_OP = 10
Const NUM_MAX_OPS_SELECIONADAS = 50
Const NUM_MAX_CATEGORIAS_SELECIONADAS = 50

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    lErro = ListaCategoria_Preenche
    If lErro <> SUCESSO Then gError 182394
    
    lErro = ComboFilial_Preenche
    If lErro <> SUCESSO Then gError 182395
    
    Call Padrao_tela

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 182394, 182395
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182396)

    End Select

    Exit Sub

End Sub

Private Sub Padrao_tela()

Dim lErro As Long

On Error GoTo Erro_Padrao_tela

    TodasFiliais.Value = vbUnchecked
    
    iFilialAnt = giFilialEmpresa

    Call Combo_Seleciona_ItemData(ComboFilial, giFilialEmpresa)

    Exit Sub

Erro_Padrao_tela:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182397)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sNomeParam As String
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim iNumOPsMarcadas As Integer
Dim iNumCatsMarcadas As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    Call Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 182398
    
    Call BotaoDesTodosOP_Click
    
    'Obtém o número de OPs que foram marcadas
    lErro = objRelOpcoes.ObterParametro("NFILIAL", sParam)
    If lErro Then gError 182399
    
    If StrParaInt(sParam) = 0 Then
        TodasFiliais.Value = vbChecked
        Call TodasFiliais_Click
    Else
        TodasFiliais.Value = vbUnchecked
        Call Combo_Seleciona_ItemData(ComboFilial, StrParaInt(sParam))
    End If
    
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 182399
    
    Call DateParaMasked(DataInicial, StrParaDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 182399

    Call DateParaMasked(DataFinal, StrParaDate(sParam))
    
    'Obtém o número de OPs que foram marcadas
    lErro = objRelOpcoes.ObterParametro("NNUMOPS", sParam)
    If lErro Then gError 182400
    
    iNumOPsMarcadas = StrParaInt(sParam)
    
    For iIndice = 1 To iNumOPsMarcadas
    
        sNomeParam = "TOP" & CStr(iIndice)

        'pega Ordem de Producao
        lErro = objRelOpcoes.ObterParametro(sNomeParam, sParam)
        If lErro Then gError 182401
        
        For iIndice2 = 0 To ListaOP.ListCount - 1
            If ListaOP.List(iIndice2) = sParam Then
                ListaOP.Selected(iIndice2) = True
            End If
        Next
        
    Next
    
    'Obtém o número de Categorias que foram marcadas
    lErro = objRelOpcoes.ObterParametro("NNUMCATS", sParam)
    If lErro Then gError 182402
    
    iNumCatsMarcadas = StrParaInt(sParam)
    
    For iIndice = 1 To iNumCatsMarcadas
    
        sNomeParam = "TCAT" & CStr(iIndice)

        'pega a Categoria
        lErro = objRelOpcoes.ObterParametro(sNomeParam, sParam)
        If lErro Then gError 182403
        
        For iIndice2 = 0 To ListaCat.ListCount - 1
            If ListaCat.List(iIndice2) = sParam Then
                ListaCat.Selected(iIndice2) = True
            End If
        Next
        
    Next

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr
    
        Case 182398 To 182403
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182404)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 182405
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 182406

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 182405
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
            
        Case 182406
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182407)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Sub Limpar_Tela()

    Call BotaoDesTodosCat_Click
    Call BotaoDesTodosOP_Click

    Call Limpa_Tela(Me)
    
    Call Padrao_tela

    ComboOpcoes.SetFocus

End Sub

Private Function Formata_E_Critica_Parametros(iFilial As Integer, sListaOPs As String, sListaCats As String, ByVal colOPs As Collection, ByVal colCats As Collection, ByVal colOrdensProducao As Collection, ByVal colCategorias As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim vValor As Variant
Dim sValor As String
Dim objOP As ClassOrdemDeProducao
Dim objCatProd As ClassCategoriaProduto

On Error GoTo Erro_Formata_E_Critica_Parametros

    For iIndice = 0 To ListaCat.ListCount - 1
        If ListaCat.Selected(iIndice) Then
            colCats.Add ListaCat.List(iIndice)
            
            Set objCatProd = New ClassCategoriaProduto
            
            objCatProd.sCategoria = Trim(ListaCat.List(iIndice))
            
            colCategorias.Add objCatProd
        End If
    Next
    
    For iIndice = 0 To ListaOP.ListCount - 1
        If ListaOP.Selected(iIndice) Then
            colOPs.Add ListaOP.List(iIndice)
            
            Set objOP = New ClassOrdemDeProducao
            
            objOP.sCodigo = Trim(Left(ListaOP.List(iIndice), TAM_COD_OP))
            objOP.iFilialEmpresa = Codigo_Extrai(Right(ListaOP.List(iIndice), Len(ListaOP.List(iIndice)) - TAM_COD_OP))
            
            colOrdensProducao.Add objOP

        End If
    Next
    
    iIndice = 0
    sListaCats = ""
    For Each vValor In colCats
        iIndice = iIndice + 1

        'Se é o primeiro
        If iIndice = 1 Then
            sListaCats = vValor
        'Se é o último
        ElseIf iIndice = colCats.Count Then
            sListaCats = sListaCats & " e " & vValor
        Else
            sListaCats = sListaCats & ", " & vValor
        End If
    Next
    
    iIndice = 0
    sListaOPs = ""
    For Each vValor In colOPs
        sValor = Trim(Left(vValor, TAM_COD_OP))
        iIndice = iIndice + 1
        If iIndice = 1 Then
            sListaOPs = sValor
        'Se é o último
        ElseIf iIndice = colOPs.Count Then
            sListaOPs = sListaOPs & " e " & sValor
        'Se é o primeiro
        Else
            sListaOPs = sListaOPs & ", " & sValor
        End If
    Next
    
    If TodasFiliais.Value = vbChecked Then
        iFilial = EMPRESA_TODA
    Else
        iFilial = ComboFilial.ItemData(ComboFilial.ListIndex)
    End If
    
    If colOPs.Count > NUM_MAX_OPS_SELECIONADAS Then gError 182444
    If colCats.Count > NUM_MAX_CATEGORIAS_SELECIONADAS Then gError 182445

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 182444
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_MAX_OP_SELECIONADAS", gErr, colOPs.Count, NUM_MAX_OPS_SELECIONADAS)

        Case 182445
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_MAX_CATEGORIAS_SELECIONADAS", gErr, colCats.Count, NUM_MAX_CATEGORIAS_SELECIONADAS)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182408)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    
    Limpar_Tela

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iIndice As Integer
Dim sListaCats As String
Dim sListaOPs As String
Dim colOPs As New Collection
Dim colCats As New Collection
Dim vValor As Variant
Dim sNomeParam As String
Dim iFilial As Integer
Dim lNumIntRel As Long
Dim colOrdensProducao As New Collection
Dim colCategorias As New Collection

On Error GoTo Erro_PreencherRelOp

    lErro = Formata_E_Critica_Parametros(iFilial, sListaOPs, sListaCats, colOPs, colCats, colOrdensProducao, colCategorias)
    If lErro <> SUCESSO Then gError 182409

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 182410

    lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(StrParaDate(DataInicial.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 182411

    lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(StrParaDate(DataFinal.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 182411

    sListaOPs = Left(sListaOPs, 80)
    lErro = objRelOpcoes.IncluirParametro("TLISTAOPS", sListaOPs)
    If lErro <> AD_BOOL_TRUE Then gError 182411

    sListaCats = Left(sListaCats, 80)
    lErro = objRelOpcoes.IncluirParametro("TLISTACATS", sListaCats)
    If lErro <> AD_BOOL_TRUE Then gError 182412
    
    lErro = objRelOpcoes.IncluirParametro("NNUMOPS", CStr(colOPs.Count))
    If lErro <> AD_BOOL_TRUE Then gError 182413
    
    lErro = objRelOpcoes.IncluirParametro("NNUMCATS", CStr(colCats.Count))
    If lErro <> AD_BOOL_TRUE Then gError 182414

    iIndice = 0
    For Each vValor In colOPs
    
        iIndice = iIndice + 1
        sNomeParam = "TOP" & CStr(iIndice)
    
        lErro = objRelOpcoes.IncluirParametro(sNomeParam, CStr(vValor))
        If lErro <> AD_BOOL_TRUE Then gError 182415
    
    Next
    
    iIndice = 0
    For Each vValor In colCats
    
        iIndice = iIndice + 1
        sNomeParam = "TCAT" & CStr(iIndice)
    
        lErro = objRelOpcoes.IncluirParametro(sNomeParam, CStr(vValor))
        If lErro <> AD_BOOL_TRUE Then gError 182416
    
    Next

    lErro = objRelOpcoes.IncluirParametro("NFILIAL", CStr(iFilial))
    If lErro <> AD_BOOL_TRUE Then gError 182417
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError 182418
    
    If bExecutando Then
    
        lErro = CF("RelOPHar_Prepara", lNumIntRel, colOrdensProducao, colCategorias)
        If lErro <> SUCESSO Then gError 182438
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 182439
    
    End If

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 182409 To 182418, 182438, 182439

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182419)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 182420

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 182421

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 182420
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 182421

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182422)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 182423

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 182423

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182424)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 182425

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 182426

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 182427

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 182425
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 182426, 182427

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182428)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182429)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_ORDEM_PRODUCAO
    Set Form_Load_Ocx = Me
    Caption = "Ordens de Produção"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpOPHar"
    
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
        
    
    End If

End Sub

Function ListaCategoria_Preenche() As Long

Dim lErro As Long
Dim colCategorias As New Collection
Dim objCatProd As ClassCategoriaProduto

On Error GoTo Erro_ListaCategoria_Preenche

    lErro = CF("CategoriasProduto_Le_Todas", colCategorias)
    If lErro <> SUCESSO And lErro <> 22542 Then gError 182430
    
    ListaCat.Clear
    
    For Each objCatProd In colCategorias
            
        ListaCat.AddItem objCatProd.sCategoria
    
    Next

    ListaCategoria_Preenche = SUCESSO

    Exit Function

Erro_ListaCategoria_Preenche:

    ListaCategoria_Preenche = gErr

    Select Case gErr
    
        Case 182430

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182431)

    End Select

    Exit Function

End Function

Function ListaOP_Preenche(ByVal iFilial As Integer, ByVal dtDataIni As Date, ByVal dtDataFim As Date) As Long

Dim lErro As Long
Dim colOP As New Collection
Dim objOP As ClassOrdemDeProducao
Dim sConteudo As String
Dim objFiliais As AdmFiliais
Dim bPreenche As Boolean

On Error GoTo Erro_ListaOP_Preenche

    lErro = CF("OrdemProducao_Le_Todas_Filial", iFilial, colOP)
    If lErro <> SUCESSO And lErro <> 22542 Then gError 182432
    
    ListaOP.Clear
    
    For Each objOP In colOP
    
        For Each objFiliais In gcolFiliais
            If objOP.iFilialEmpresa = objFiliais.iCodFilial Then
                Exit For
            End If
        Next
    
        sConteudo = objOP.sCodigo & String(TAM_COD_OP - Len(objOP.sCodigo), " ") & CStr(objOP.iFilialEmpresa) & SEPARADOR & objFiliais.sNome
        
        bPreenche = True
        If dtDataIni <> DATA_NULA Then
            If objOP.dtDataEmissao < dtDataIni Then bPreenche = False
        End If
        If dtDataFim <> DATA_NULA Then
            If objOP.dtDataEmissao > dtDataFim Then bPreenche = False
        End If
        
        If bPreenche Then
            ListaOP.AddItem sConteudo
        End If
    
    Next

    ListaOP_Preenche = SUCESSO

    Exit Function

Erro_ListaOP_Preenche:

    ListaOP_Preenche = gErr

    Select Case gErr
    
        Case 182432

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182433)

    End Select

    Exit Function

End Function

Function ComboFilial_Preenche() As Long

Dim lErro As Long
Dim objFiliais As AdmFiliais

On Error GoTo Erro_ComboFilial_Preenche

    For Each objFiliais In gcolFiliais

        If objFiliais.iCodFilial <> EMPRESA_TODA And objFiliais.iCodFilial > DELTA_FILIALREAL_OFICIAL Then
        
            ComboFilial.AddItem CStr(objFiliais.iCodFilial) & SEPARADOR & objFiliais.sNome
            ComboFilial.ItemData(ComboFilial.NewIndex) = objFiliais.iCodFilial
    
        End If
        
    Next
    
    ComboFilial_Preenche = SUCESSO

    Exit Function

Erro_ComboFilial_Preenche:

    ComboFilial_Preenche = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182435)

    End Select

    Exit Function

End Function

Function ComboFilial_AtualizaOP() As Long

Dim lErro As Long
Dim iFilial As Integer

On Error GoTo Erro_ComboFilial_AtualizaOP

    If ComboFilial.ListIndex = -1 Then
        iFilial = 0
    Else
        iFilial = ComboFilial.ItemData(ComboFilial.ListIndex)
    End If
    
    If iFilial <> iFilialAnt Or dtDataInicialAnt <> StrParaDate(DataInicial.Text) Or dtDataFinalAnt <> StrParaDate(DataFinal.Text) Then
    
        lErro = ListaOP_Preenche(iFilial, StrParaDate(DataInicial.Text), StrParaDate(DataFinal.Text))
        If lErro <> SUCESSO Then gError 182436
        
        iFilialAnt = iFilial
        dtDataInicialAnt = StrParaDate(DataInicial.Text)
        dtDataFinalAnt = StrParaDate(DataFinal.Text)
        
    End If
    
    ComboFilial_AtualizaOP = SUCESSO

    Exit Function

Erro_ComboFilial_AtualizaOP:

    ComboFilial_AtualizaOP = gErr

    Select Case gErr
    
        Case 182436

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182437)

    End Select

    Exit Function

End Function

Private Sub TodasFiliais_Click()
    If TodasFiliais.Value = vbChecked Then
        ComboFilial.ListIndex = -1
        ComboFilial.Enabled = False
    Else
        ComboFilial.Enabled = True
        Call Combo_Seleciona_ItemData(ComboFilial, giFilialEmpresa)
    End If
End Sub

Private Sub TodasFiliais_Change()
    If TodasFiliais.Value = vbChecked Then
        ComboFilial.ListIndex = -1
        ComboFilial.Enabled = False
    Else
        ComboFilial.Enabled = True
        Call Combo_Seleciona_ItemData(ComboFilial, giFilialEmpresa)
    End If
End Sub

Private Sub ComboFilial_Change()
    Call ComboFilial_AtualizaOP
End Sub

Private Sub ComboFilial_Click()
    Call ComboFilial_AtualizaOP
End Sub

Private Sub BotaoDesTodosOP_Click()

Dim iIndice As Integer

    For iIndice = 0 To ListaOP.ListCount - 1
        ListaOP.Selected(iIndice) = False
    Next
    
End Sub

Private Sub BotaoMarTodosOP_Click()

Dim iIndice As Integer

    For iIndice = 0 To ListaOP.ListCount - 1
        ListaOP.Selected(iIndice) = True
    Next

End Sub

Private Sub BotaoDesTodosCat_Click()

Dim iIndice As Integer

    For iIndice = 0 To ListaCat.ListCount - 1
        ListaCat.Selected(iIndice) = False
    Next
    
End Sub

Private Sub BotaoMarTodosCat_Click()

Dim iIndice As Integer

    For iIndice = 0 To ListaCat.ListCount - 1
        ListaCat.Selected(iIndice) = True
    Next

End Sub

Private Sub DataFinal_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataFinal)
End Sub

Private Sub DataInicial_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataInicial)
End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then gError 182543

    End If

    Call ComboFilial_AtualizaOP

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 182543

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182544)

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
        If lErro <> SUCESSO Then gError 182545

    End If

    Call ComboFilial_AtualizaOP

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 182545

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182546)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 182547

    Call DataInicial_Validate(bSGECancelDummy)

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 182547
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182548)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 182549

    Call DataInicial_Validate(bSGECancelDummy)

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 182549
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182550)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 182551

    Call DataFinal_Validate(bSGECancelDummy)

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 182551
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182552)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 182553
    
    Call DataFinal_Validate(bSGECancelDummy)

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 182553
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182554)

    End Select

    Exit Sub

End Sub


