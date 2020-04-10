VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Begin VB.UserControl RelOpNecessRoteiroOcx 
   ClientHeight    =   3660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7830
   KeyPreview      =   -1  'True
   ScaleHeight     =   3660
   ScaleWidth      =   7830
   Begin VB.CommandButton BotaoKits 
      Caption         =   "&Kits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   1410
      Picture         =   "RelOpNecessidadeRoteiro.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Visualiza os Kits cadastrados"
      Top             =   3000
      Width           =   1200
   End
   Begin VB.CheckBox Recursivo 
      Caption         =   "Considerar os Roteiros de Fabricação dos Produtos Intermediários"
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
      Left            =   450
      TabIndex        =   5
      Top             =   2610
      Width           =   6165
   End
   Begin VB.CommandButton BotaoVerRoteiros 
      Caption         =   "&Roteiros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   120
      Picture         =   "RelOpNecessidadeRoteiro.ctx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Roteiros de Fabricação cadastrados"
      Top             =   3000
      Width           =   1200
   End
   Begin VB.ComboBox UM 
      Height          =   315
      Left            =   1725
      TabIndex        =   3
      Top             =   1995
      Width           =   1125
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpNecessidadeRoteiro.ctx":0614
      Left            =   840
      List            =   "RelOpNecessidadeRoteiro.ctx":0616
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
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
      Left            =   3960
      Picture         =   "RelOpNecessidadeRoteiro.ctx":0618
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5565
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1620
         Picture         =   "RelOpNecessidadeRoteiro.ctx":071A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1115
         Picture         =   "RelOpNecessidadeRoteiro.ctx":0898
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   610
         Picture         =   "RelOpNecessidadeRoteiro.ctx":0DCA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "RelOpNecessidadeRoteiro.ctx":0F54
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Produto 
      Height          =   315
      Left            =   1725
      TabIndex        =   1
      Top             =   960
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Versao 
      Height          =   315
      Left            =   1725
      TabIndex        =   2
      Top             =   1485
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Quantidade 
      Height          =   315
      Left            =   4440
      TabIndex        =   4
      Top             =   1995
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label Label3 
      Caption         =   "UM:"
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
      Left            =   975
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   18
      Top             =   2070
      Width           =   405
   End
   Begin VB.Label Label2 
      Caption         =   "Quantidade:"
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
      Left            =   3285
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   17
      Top             =   2070
      Width           =   1110
   End
   Begin VB.Label LabelVersao 
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
      Height          =   255
      Left            =   675
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   16
      Top             =   1530
      Width           =   690
   End
   Begin VB.Label LabelProduto 
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
      Height          =   255
      Left            =   570
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   15
      Top             =   1005
      Width           =   855
   End
   Begin VB.Label DescProd 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3345
      TabIndex        =   14
      Top             =   960
      Width           =   4350
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
      TabIndex        =   13
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpNecessRoteiroOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoRoteiroDeFabricacao As AdmEvento
Attribute objEventoRoteiroDeFabricacao.VB_VarHelpID = -1
Private WithEvents objEventoKit As AdmEvento
Attribute objEventoKit.VB_VarHelpID = -1
Private WithEvents objEventoVersao As AdmEvento
Attribute objEventoVersao.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio
Dim giFocoInicial As Integer

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoProduto = Nothing
    Set objEventoRoteiroDeFabricacao = Nothing
    Set objEventoVersao = Nothing
    Set objEventoKit = Nothing
    
End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 141681

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 141682

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, Produto, DescProd)
    If lErro <> SUCESSO Then gError 141683

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 141681, 141683
            'erro tratado na rotina chamada

        Case 141682
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170164)

    End Select

    Exit Sub

End Sub

Private Sub LabelProduto_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProduto_Click

    'Verifica se o produto foi preenchido
    If Len(Produto.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 141684

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_LabelProduto_Click:

    Select Case gErr

        Case 141684
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170165)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 141685
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 141686
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 141685
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
            
        Case 141686
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170166)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd As String) As Long
'monta a expressão de seleção
'recebe os produtos inicial e final no formato do BD

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

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170167)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sProd As String) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se o produto inicial é maior que o produto final

Dim iProdPreenchido As Integer
Dim lErro As Long
Dim objRoteiro As New ClassRoteirosDeFabricacao

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata o Produto Inicial
    lErro = CF("Produto_Formata", Produto.Text, sProd, iProdPreenchido)
    If lErro <> SUCESSO Then gError 141687
    
    If iProdPreenchido <> PRODUTO_PREENCHIDO Then sProd = ""
    
    If Len(Trim(Versao.Text)) = 0 Then gError 141688
    
    If StrParaDbl(Quantidade.Text) < QTDE_ESTOQUE_DELTA Then gError 141689
    
    objRoteiro.sProdutoRaiz = sProd
    objRoteiro.sVersao = Versao.Text
    
'    lErro = CF("RoteirosDeFabricacao_Le", objRoteiro)
'    If lErro <> SUCESSO And lErro <> 134617 Then gError 141718
'
'    If lErro <> SUCESSO Then gError 141719

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 141687
            Produto.SetFocus
            
        Case 141688
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAO_NAO_PREENCHIDA", gErr)
            Versao.SetFocus
        
        Case 141689
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr)
            Quantidade.SetFocus
            
        Case 141718
        
        Case 141719
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTEIROSDEFABRICACAO_NAO_CADASTRADO", gErr, objRoteiro.sProdutoRaiz, objRoteiro.sVersao)
           
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170168)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutar As Boolean = False, Optional iMaiorNivel As Integer) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd As String
Dim lNumIntRel As Long
Dim objRoteiroNecess As New ClassRoteiroNecessidade
Dim bRecursivo As Boolean

On Error GoTo Erro_PreencherRelOp
    
    lErro = Formata_E_Critica_Parametros(sProd)
    If lErro <> SUCESSO Then gError 141690

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 141691

    lErro = objRelOpcoes.IncluirParametro("TPROD", sProd)
    If lErro <> AD_BOOL_TRUE Then gError 141692

    lErro = objRelOpcoes.IncluirParametro("TVERSAO", Versao.Text)
    If lErro <> AD_BOOL_TRUE Then gError 141693

    lErro = objRelOpcoes.IncluirParametro("TUM", UM.Text)
    If lErro <> AD_BOOL_TRUE Then gError 141694

    lErro = objRelOpcoes.IncluirParametro("NQTDE", Quantidade.Text)
    If lErro <> AD_BOOL_TRUE Then gError 141695

    lErro = objRelOpcoes.IncluirParametro("NRECURSIVO", CStr(Recursivo.Value))
    If lErro <> AD_BOOL_TRUE Then gError 141695

    If bExecutar Then
    
        objRoteiroNecess.dQuantidade = StrParaDbl(Quantidade.Text)
        objRoteiroNecess.sProdutoRaiz = sProd
        objRoteiroNecess.sUM = UM.Text
        objRoteiroNecess.sVersao = Versao.Text
        
        If Recursivo.Value = vbChecked Then
            bRecursivo = True
        Else
            bRecursivo = False
        End If

        lErro = CF("RelNecessidadeRoteiro_Prepara", lNumIntRel, objRoteiroNecess, bRecursivo)
        If lErro <> SUCESSO Then gError 141696
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 141697
    
    End If

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd)
    If lErro <> SUCESSO Then gError 141698

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 141690 To 141697
            'erro tratado nas rotinas chamadas
            
        Case 141698
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_SEM_DADOS", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170169)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 141699

    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPROD", sParam)
    If lErro Then gError 141700

    lErro = CF("Traz_Produto_MaskEd", sParam, Produto, DescProd)
    If lErro <> SUCESSO Then gError 134573

    'pega Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TVERSAO", sParam)
    If lErro Then gError 141701
    
    Versao.Text = sParam

    'pega Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TUM", sParam)
    If lErro Then gError 141702
    
    UM.Text = sParam

    'pega Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("NQTDE", sParam)
    If lErro Then gError 141703
    
    Quantidade.Text = Formata_Estoque(StrParaDbl(sParam))
    
    Call Produto_Validate(bSGECancelDummy)

    'pega Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("NRECURSIVO", sParam)
    If lErro Then gError 141702
    
    Recursivo.Value = StrParaInt(sParam)

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 141699 To 141703

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170170)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 141704

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 141705

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 141704
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 141705
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170171)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim iMaiorNivel As Integer

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 141706

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 141706
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170172)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 141707

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 141708

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 141709

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 141707
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 141708, 141709
            'erro tratado nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170173)

    End Select

    Exit Sub

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)
    DescProd.Caption = ""
    
    UM.Clear
    
    Recursivo.Value = vbUnchecked

    ComboOpcoes.SetFocus

End Sub

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

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim objKit As New ClassKit
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Produto_Validate

    sProduto = Produto.Text

    'Critica o formato do Produto e se existe no BD
    lErro = CF("Produto_Critica2", sProduto, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then Error 141710

    'se o produto não estiver cadastrado ==> erro
    If lErro = 25041 Then gError 141711
       
    DescProd.Caption = objProduto.sDescricao

    lErro = CarregaComboUM(objProduto)
    If lErro <> SUCESSO Then gError 141714
    
    objKit.sProdutoRaiz = objProduto.sCodigo

    lErro = CF("Kit_Le_Padrao", objKit)
    If lErro <> SUCESSO And lErro <> 106304 Then gError 141713

    Versao.Text = objKit.sVersao
  
    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 141710, 141712, 141713, 141714
            'erro tratado na rotina chamada

        Case 141711
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170174)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_RelOpProdutos_Form_Load
      
    Set objEventoProduto = New AdmEvento
    Set objEventoRoteiroDeFabricacao = New AdmEvento
    Set objEventoVersao = New AdmEvento
    Set objEventoKit = New AdmEvento

    'inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 141715

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_RelOpProdutos_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 141715
            'erro tratado na rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170175)

    End Select
   
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_PRODUTOS
    Set Form_Load_Ocx = Me
    Caption = "Necessidades de Produção com o Roteiro de Fabricação"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpNecessidadeRoteiro"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is Produto Then
            Call LabelProduto_Click
        ElseIf Me.ActiveControl Is Versao Then
            Call LabelVersao_Click
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
Private Sub LabelProduto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProduto, Source, X, Y)
End Sub

Private Sub LabelProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProduto, Button, Shift, X, Y)
End Sub

Private Sub DescProd_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescProd, Source, X, Y)
End Sub

Private Sub DescProd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescProd, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub BotaoVerRoteiros_Click()

Dim lErro As Long
Dim objRoteirosDeFabricacao As New ClassRoteirosDeFabricacao
Dim sProdutoFormatado As String
Dim colSelecao As New Collection
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoVerRoteiros_Click

    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 141716

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""

    objRoteirosDeFabricacao.sProdutoRaiz = sProdutoFormatado
    
    If Len(Trim(Versao.Text)) <> 0 Then
        objRoteirosDeFabricacao.sVersao = Versao.Text
    End If

    Call Chama_Tela("RoteirosDeFabricacaoLista", colSelecao, objRoteirosDeFabricacao, objEventoRoteiroDeFabricacao)

    Exit Sub

Erro_BotaoVerRoteiros_Click:

    Select Case gErr

        Case 141716

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170176)

    End Select

    Exit Sub

End Sub

Private Sub objEventoRoteiroDeFabricacao_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRoteirosDeFabricacao As ClassRoteirosDeFabricacao
Dim sProdMask As String

On Error GoTo Erro_objEventoRoteiroDeFabricacao_evSelecao

    Set objRoteirosDeFabricacao = obj1
    
    lErro = Mascara_RetornaProdutoTela(objRoteirosDeFabricacao.sProdutoRaiz, sProdMask)
    If lErro <> SUCESSO Then gError 141716

    Produto.PromptInclude = False
    Produto.Text = sProdMask
    Produto.PromptInclude = True

    Call Produto_Validate(bSGECancelDummy)

    Quantidade.Text = Formata_Estoque(objRoteirosDeFabricacao.dQuantidade)
    UM.Text = objRoteirosDeFabricacao.sUM
    Versao.Text = objRoteirosDeFabricacao.sVersao

    Me.Show

    Exit Sub

Erro_objEventoRoteiroDeFabricacao_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170177)

    End Select

    Exit Sub

End Sub

Private Function CarregaComboUM(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim objClasseUM As ClassClasseUM
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim colSiglas As New Collection
Dim sUnidadeMed As String
Dim iIndice As Integer

On Error GoTo Erro_CarregaComboUM

    Set objClasseUM = New ClassClasseUM
    
    objClasseUM.iClasse = objProduto.iClasseUM
    
    'Preenche a List da Combo UnidadeMed com as UM's da Competencia
    lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
    If lErro <> SUCESSO And lErro <> 22539 Then gError 141717

    'Se tem algum valor para UM na Tela
    If Len(UM.Text) > 0 Then
        'Guardo o valor da UM da Tela
        sUnidadeMed = UM.Text
    Else
        'Senão coloco a do Estoque do Produto
        sUnidadeMed = objProduto.sSiglaUMEstoque
    End If
    
    'Limpar as Unidades utilizadas anteriormente
    UM.Clear

    For Each objUnidadeDeMedida In colSiglas
        UM.AddItem objUnidadeDeMedida.sSigla
    Next

    UM.AddItem ""

    'Tento selecionar na Combo a Unidade anterior
    If UM.ListCount <> 0 Then

        For iIndice = 0 To UM.ListCount - 1

            If UM.List(iIndice) = sUnidadeMed Then
                UM.ListIndex = iIndice
                Exit For
            End If
        Next
    End If
    
    CarregaComboUM = SUCESSO
    
    Exit Function

Erro_CarregaComboUM:

    CarregaComboUM = gErr

    Select Case gErr

        Case 141717

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170178)

    End Select

    Exit Function

End Function

Private Sub objEventoVersao_evSelecao(obj1 As Object)

Dim objKit As ClassKit
Dim lErro As Long

On Error GoTo Erro_objEventoVersao_evSelecao

    Set objKit = obj1

    Versao.Text = objKit.sVersao
        
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoVersao_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174244)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelVersao_Click()

Dim lErro As Long
Dim objKit As New ClassKit
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_LabelVersao_Click

    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134756

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        objKit.sProdutoRaiz = sProdutoFormatado
        If Len(Trim(Versao.ClipText)) > 0 Then objKit.sVersao = Versao.Text
            
        colSelecao.Add sProdutoFormatado
        
        Call Chama_Tela("KitVersaoLista", colSelecao, objKit, objEventoVersao)
    
    Else
         gError 134757
         
    End If

    Exit Sub

Erro_LabelVersao_Click:

    Select Case gErr

        Case 134756
        
        Case 134757
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTORAIZKIT_NAO_PREENCHIDO2", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174272)

    End Select

    Exit Sub

End Sub

Private Sub BotaoKits_Click()

Dim lErro As Long
Dim objKit As New ClassKit
Dim sProdutoFormatado As String
Dim colSelecao As New Collection
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoKits_Click

    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 141716

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""

    objKit.sProdutoRaiz = sProdutoFormatado
    
    If Len(Trim(Versao.Text)) <> 0 Then
        objKit.sVersao = Versao.Text
    End If

    Call Chama_Tela("KitLista", colSelecao, objKit, objEventoKit)

    Exit Sub

Erro_BotaoKits_Click:

    Select Case gErr

        Case 141716

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170176)

    End Select

    Exit Sub

End Sub

Private Sub objEventoKit_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objKit As ClassKit
Dim sProdMask As String

On Error GoTo Erro_objEventoKit_evSelecao

    Set objKit = obj1
    
    lErro = Mascara_RetornaProdutoTela(objKit.sProdutoRaiz, sProdMask)
    If lErro <> SUCESSO Then gError 141716

    Produto.PromptInclude = False
    Produto.Text = sProdMask
    Produto.PromptInclude = True

    Call Produto_Validate(bSGECancelDummy)

    Versao.Text = objKit.sVersao

    Me.Show

    Exit Sub

Erro_objEventoKit_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170177)

    End Select

    Exit Sub

End Sub
