VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Begin VB.UserControl ServItemServ 
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7530
   KeyPreview      =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   7530
   Begin VB.CommandButton BotaoServicos 
      Caption         =   "Serviços"
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
      Left            =   3300
      TabIndex        =   15
      Top             =   300
      Width           =   1605
   End
   Begin VB.CommandButton BotaoParaCima 
      Height          =   390
      Left            =   7020
      Picture         =   "ServItemServGR.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2265
      Width           =   390
   End
   Begin VB.CommandButton BotaoParaBaixo 
      Height          =   390
      Left            =   7020
      Picture         =   "ServItemServGR.ctx":01C2
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2910
      Width           =   390
   End
   Begin VB.CommandButton BotaoItensServico 
      Caption         =   "Itens de Serviço"
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
      Left            =   4860
      TabIndex        =   7
      Top             =   4365
      Width           =   1965
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5250
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   165
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "ServItemServGR.ctx":0384
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ServItemServGR.ctx":04DE
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ServItemServGR.ctx":0668
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ServItemServGR.ctx":0B9A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.TextBox TextDescItemServico 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      HideSelection   =   0   'False
      Left            =   1770
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1665
      Width           =   4575
   End
   Begin MSMask.MaskEdBox MaskItemServico 
      Height          =   225
      Left            =   330
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid GridItensServico 
      Height          =   2910
      Left            =   180
      TabIndex        =   1
      Top             =   1320
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   5133
      _Version        =   393216
      Rows            =   12
      ForeColorSel    =   16777215
      AllowBigSelection=   0   'False
      FocusRect       =   2
      SelectionMode   =   1
   End
   Begin MSMask.MaskEdBox MaskServico 
      Height          =   315
      Left            =   1290
      TabIndex        =   0
      Top             =   315
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label Label9 
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
      Height          =   165
      Left            =   255
      TabIndex        =   14
      Top             =   930
      Width           =   930
   End
   Begin VB.Label LabelDescricaoServico 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1275
      TabIndex        =   13
      Top             =   885
      Width           =   5565
   End
   Begin VB.Label LabelServico 
      AutoSize        =   -1  'True
      Caption         =   "Serviço:"
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
      Left            =   450
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   12
      Top             =   360
      Width           =   720
   End
End
Attribute VB_Name = "ServItemServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis que serão utilizadas pelo grid
Dim objGridItensServico As AdmGrid
Dim iGrid_ItemServico_Col As Integer
Dim iGrid_Decricao_Col As Integer

Dim iAlterado As Integer

'Variáveis que serão utilizadas pelo browser
Private WithEvents objEventoItemServico As AdmEvento
Attribute objEventoItemServico.VB_VarHelpID = -1
Private WithEvents objEventoServicoItemServico As AdmEvento
Attribute objEventoServicoItemServico.VB_VarHelpID = -1
Private WithEvents objEventoServico As AdmEvento
Attribute objEventoServico.VB_VarHelpID = -1

'Definições das Constantes
Private Const NUM_MAX_ITEM_SERVICOS = 100

Function Inicializa_Grid_ItensServico(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Item Serviço")
    objGridInt.colColuna.Add ("Descrição")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (MaskItemServico.Name)
    objGridInt.colCampo.Add (TextDescItemServico.Name)

    'Colunas do Grid
    iGrid_ItemServico_Col = 1
    iGrid_Decricao_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridItensServico

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 11

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITEM_SERVICOS + 1

    'Largura da primeira coluna
    GridItensServico.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Call Reconfigura_Linha_Grid

    Inicializa_Grid_ItensServico = SUCESSO

End Function

Function Trata_Parametros(Optional objProduto As ClassProduto) As Long

Dim lErro As Long
Dim objServItemServ As New ClassServItemServ

On Error GoTo Erro_Trata_Parametros

    'Verifica se houve passagem de parametro
    If Not (objProduto Is Nothing) Then
    
        objServItemServ.sProduto = objProduto.sCodigo

        'Traz os itens de serviço associados ao serviço
        lErro = Traz_ServItemServ_Tela(objServItemServ)
        If lErro <> SUCESSO And lErro <> 97546 Then gError 95484
        
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objGridItensServico = New AdmGrid

    'Inicializa o grid de Itens Serviços
    lErro = Inicializa_Grid_ItensServico(objGridItensServico)
    If lErro <> SUCESSO Then gError 95483

    Set objEventoItemServico = New AdmEvento
    Set objEventoServico = New AdmEvento
    Set objEventoServicoItemServico = New AdmEvento

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", MaskServico)
    If lErro <> SUCESSO Then Error 95482

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 95482, 95483

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objServico As New ClassServico

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Servicos"

    'Move os dados da Tela para o Obj
    lErro = Move_Tela_Memoria(objServico)
    If lErro <> SUCESSO Then gError 95485

    'Preenche a coleção colCampoValor
    colCampoValor.Add "Codigo", objServico.sProduto, STRING_PRODUTO, "Codigo"
    
    'Adiciona o filtro por filial
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 95485

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()
'Exclui a associacao de Servico com os itens de servico

Dim lErro As Long
Dim objServItemServ As New ClassServItemServ
Dim iProdutoPreenchido As Integer
Dim sProduto As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Botao_Excluir_Click
    
    'Coloca o cursor com formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o Servico está preenchido
    If Len(Trim(MaskServico.ClipText)) = 0 Then gError 97523
    
    'Guarda o Codigo do servico no objServItemServ
    lErro = CF("Produto_Formata", MaskServico.Text, sProduto, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 97524
    
    objServItemServ.sProduto = sProduto
    
    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_SERVICOITEMSERVICO", objServItemServ.sProduto)

    'Se a resposta for positiva
    If vbMsgRes = vbYes Then
        
        'Exclui a associacao de Servico X Item de Servico
        lErro = CF("ServicoItemServico_Exclui", objServItemServ)
        If lErro <> SUCESSO Then gError 97525
        
        'Fecha o Comando de Setas
        Call ComandoSeta_Fechar(Me.Name)
        
        'Limpa a tela
        Call Limpa_Tela_ServItemServ
        
        iAlterado = 0
    
    End If
    
    'Retorna o cursor ao formato default
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_Botao_Excluir_Click:
    
    Select Case gErr
    
        Case 97523
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PREENCHIDO", gErr)
        
        Case 97524, 97525
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select
    
    'Fecha o Comando de Setas
    Call ComandoSeta_Fechar(Me.Name)
    
    'Retorna o cursor ao formato default
    GL_objMDIForm.MousePointer = vbDefault

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Controla toda a rotina de gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 97506

    'Limpa a Tela
    Call Limpa_Tela_ServItemServ

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 97506

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

     End Select

     Exit Sub

End Sub

Private Sub Limpa_Tela_ServItemServ()

    Call Grid_Limpa(objGridItensServico)

    Call Limpa_Tela(Me)

    'Limpa a descricao do serviço
    LabelDescricaoServico.Caption = ""

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'Verifica se existe algo para ser salvo antes de limpar a tela
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> AD_SQL_SUCESSO Then gError 97511

    'Limpa a Tela
    Call Limpa_Tela_ServItemServ

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 97511

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

     End Select

     Exit Sub


End Sub

Private Sub BotaoParaCima_Click()

Dim lErro As Long
Dim sItem_Acima As String
Dim sDesc_Acima As String
Dim sItem As String
Dim sDesc As String

On Error GoTo Erro_BotaoParaCima_Click

    'Se não tem linha selecionada => Erro
    If GridItensServico.Row = 0 Then gError 97504

    'Se nao for a primeira linha...
    If GridItensServico.Row > 1 Then

        If Len(Trim(GridItensServico.TextMatrix(GridItensServico.Row, iGrid_ItemServico_Col))) > 0 Then

            'Guarda o conteúdo da linha selecionada e a sua superior
            sItem = GridItensServico.TextMatrix(GridItensServico.Row, iGrid_ItemServico_Col)
            sDesc = GridItensServico.TextMatrix(GridItensServico.Row, iGrid_Decricao_Col)
            sItem_Acima = GridItensServico.TextMatrix(GridItensServico.Row - 1, iGrid_ItemServico_Col)
            sDesc_Acima = GridItensServico.TextMatrix(GridItensServico.Row - 1, iGrid_Decricao_Col)

            'Troca o conteudo da mascara
            MaskItemServico.PromptInclude = False
            MaskItemServico.Text = sItem_Acima
            MaskItemServico.PromptInclude = False

            'Troca o Conteúdo
            GridItensServico.TextMatrix(GridItensServico.Row, iGrid_ItemServico_Col) = sItem_Acima
            GridItensServico.TextMatrix(GridItensServico.Row, iGrid_Decricao_Col) = sDesc_Acima
            GridItensServico.TextMatrix(GridItensServico.Row - 1, iGrid_ItemServico_Col) = sItem
            GridItensServico.TextMatrix(GridItensServico.Row - 1, iGrid_Decricao_Col) = sDesc

            GridItensServico.Row = GridItensServico.Row - 1
            GridItensServico.RowSel = GridItensServico.Row
            GridItensServico.ColSel = iGrid_Decricao_Col

        End If

    End If

    Exit Sub

Erro_BotaoParaCima_Click:

    Select Case gErr

        Case 97504
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub


Private Sub BotaoParaBaixo_Click()

Dim lErro As Long
Dim sItem_Abaixo As String
Dim sDesc_Abaixo As String
Dim sItem As String
Dim sDesc As String

On Error GoTo Erro_BotaoParaBaixo_Click

    'Se não tem linha selecionada => Erro
    If GridItensServico.Row = 0 Then gError 97505

    'Se nao for a última linha...
    If GridItensServico.Row < objGridItensServico.iLinhasExistentes Then

        If Len(Trim(GridItensServico.TextMatrix(GridItensServico.Row + 1, iGrid_ItemServico_Col))) > 0 Then
            
            'Guarda o conteúdo da linha selecionada e a sua antecessora
            sItem = GridItensServico.TextMatrix(GridItensServico.Row, iGrid_ItemServico_Col)
            sDesc = GridItensServico.TextMatrix(GridItensServico.Row, iGrid_Decricao_Col)
            sItem_Abaixo = GridItensServico.TextMatrix(GridItensServico.Row + 1, iGrid_ItemServico_Col)
            sDesc_Abaixo = GridItensServico.TextMatrix(GridItensServico.Row + 1, iGrid_Decricao_Col)
    
            'Troca o conteudo da mascara
            MaskItemServico.PromptInclude = False
            MaskItemServico.Text = sItem_Abaixo
            MaskItemServico.PromptInclude = False
    
            'Troca o Conteúdo
            GridItensServico.TextMatrix(GridItensServico.Row, iGrid_ItemServico_Col) = sItem_Abaixo
            GridItensServico.TextMatrix(GridItensServico.Row, iGrid_Decricao_Col) = sDesc_Abaixo
            GridItensServico.TextMatrix(GridItensServico.Row + 1, iGrid_ItemServico_Col) = sItem
            GridItensServico.TextMatrix(GridItensServico.Row + 1, iGrid_Decricao_Col) = sDesc
    
            GridItensServico.Row = GridItensServico.Row + 1
            GridItensServico.RowSel = GridItensServico.Row
            GridItensServico.ColSel = iGrid_Decricao_Col

        End If
        
    End If

    Exit Sub

Erro_BotaoParaBaixo_Click:

    Select Case gErr

        Case 97505
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub LabelServico_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelServico_Click

    'Verifica se o serviço foi preenchido
    If Len(Trim(MaskServico.ClipText)) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", MaskServico.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 95491

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ServItemServLista", colSelecao, objProduto, objEventoServicoItemServico)

    Exit Sub

Erro_LabelServico_Click:

    Select Case gErr

        Case 95491

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub maskServico_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim vbMsgRes As VbMsgBoxResult
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_MaskServico_Validate

    'Se o Serviço está Preenchido...
    If Len(Trim(MaskServico.ClipText)) <> 0 Then

        lErro = CF("Produto_Critica_Filial", MaskServico.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 51381 Then gError 95487

        If lErro = 51381 Then gError 95489

        If objProduto.iFaturamento = PRODUTO_NAO_VENDAVEL Then gError 95490

        'Preenche ProdutoDescricao com Descrição do Produto
        LabelDescricaoServico.Caption = objProduto.sDescricao

    Else

        'Limpa a descricao
        LabelDescricaoServico.Caption = ""

    End If

    Exit Sub

Erro_MaskServico_Validate:

    Cancel = True

    Select Case gErr

        Case 95487

        Case 95489 'Não encontrou Produto no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_SERVICO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Produto", objProduto)

            Else
                'Limpa DescricaoProduto
                LabelDescricaoServico.Caption = ""

            End If

        Case 95490
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PODE_SER_VENDIDO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub BotaoServicos_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoServicos_Click

    'Verifica se o serviço foi preenchido
    If Len(Trim(MaskServico.ClipText)) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", MaskServico.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 95491

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoServico)

    Exit Sub

Erro_BotaoServicos_Click:

    Select Case gErr

        Case 95491

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub objEventoItemServico_evSelecao(obj1 As Object)

Dim objItemServico As New ClassItemServico
Dim sItemServico As String
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_objEventoItemServico_evSelecao

    Set objItemServico = obj1

    'Verifica se alguma linha está selecionada
    If GridItensServico.Row < 1 Then Exit Sub
    
    MaskItemServico.Text = CStr(objItemServico.iCodigo)
    
    'Verificar se alguma linha do grid já contém o codigo selecionado no browser
    For iIndice = 1 To objGridItensServico.iLinhasExistentes
        If iIndice <> GridItensServico.Row Then
            If GridItensServico.TextMatrix(iIndice, iGrid_ItemServico_Col) = MaskItemServico.Text Then gError 97556
        End If
    Next

    GridItensServico.TextMatrix(GridItensServico.Row, iGrid_ItemServico_Col) = objItemServico.iCodigo
    GridItensServico.TextMatrix(GridItensServico.Row, iGrid_Decricao_Col) = objItemServico.sDescricao
    
    'Alteracao feita por Daniel em 15/02/2002
    'Incrementa o número dde linhas do grid
    objGridItensServico.iLinhasExistentes = objGridItensServico.iLinhasExistentes + 1
    
    Me.Show

    Exit Sub

Erro_objEventoItemServico_evSelecao:

    Select Case gErr

        Case 97556
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_SERVICO_JA_EXISTENTE_GRIDITENSSERVICO", gErr, objItemServico.iCodigo, iIndice)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha o Comando de Setas
    Call ComandoSeta_Fechar(Me.Name)

    Exit Sub

End Sub

Private Sub objEventoServico_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim bCancel As Boolean

On Error GoTo Erro_objEventoServico_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 95492

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 95493

    'Traz para a tela o servico e a descricao
    MaskServico.PromptInclude = False
    MaskServico.Text = objProduto.sCodigo
    MaskServico.PromptInclude = True
    LabelDescricaoServico.Caption = objProduto.sDescricao
    
    Me.Show

    Exit Sub

Erro_objEventoServico_evSelecao:

    Select Case gErr

        Case 95492

        Case 95493
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha o Comando de Setas
    Call ComandoSeta_Fechar(Me.Name)

    Exit Sub

End Sub

Private Sub GridItensServico_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItensServico, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensServico, iAlterado)
    End If

End Sub

Private Sub GridItensServico_GotFocus()

    Call Grid_Recebe_Foco(objGridItensServico)

End Sub

Private Sub GridItensServico_EnterCell()

    Call Grid_Entrada_Celula(objGridItensServico, iAlterado)

End Sub

Private Sub GridItensServico_LeaveCell()

    Call Saida_Celula(objGridItensServico)

End Sub

Private Sub GridItensServico_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridItensServico)

End Sub

Private Sub GridItensServico_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItensServico, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensServico, iAlterado)
    End If

End Sub

Private Sub GridItensServico_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItensServico)

End Sub

Private Sub GridItensServico_RowColChange()

    Call Grid_RowColChange(objGridItensServico)

End Sub

Private Sub GridItensServico_Scroll()

    Call Grid_Scroll(objGridItensServico)

End Sub

Private Sub MaskItemServico_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskItemServico_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensServico)

End Sub

Private Sub MaskItemServico_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensServico)

End Sub

Private Sub MaskItemServico_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensServico.objControle = MaskItemServico
    lErro = Grid_Campo_Libera_Foco(objGridItensServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        lErro = Saida_Celula_MaskItemServico(objGridInt)
        If lErro <> SUCESSO Then gError 95493

    End If

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 95494

    iAlterado = REGISTRO_ALTERADO

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 95494, 95493

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MaskItemServico(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItemServico As New ClassItemServico

On Error GoTo Erro_Saida_Celula_MaskItemServico

    Set objGridInt.objControle = MaskItemServico

    'Verifica se o Item está preenchido
    If Len(Trim(MaskItemServico.Text)) > 0 Then

        objItemServico.iCodigo = CInt(MaskItemServico.Text)

        'Verifica se já existe a sigla em outra linha do Grid
        For iIndice = 1 To objGridItensServico.iLinhasExistentes
            If iIndice <> GridItensServico.Row Then
                If GridItensServico.TextMatrix(iIndice, iGrid_ItemServico_Col) = MaskItemServico.Text Then gError 95495
            End If
        Next

        'Verifica se o Item com o codigo em questao existe
        lErro = CF("ItemServico_Le", objItemServico)
        If lErro <> SUCESSO And lErro <> 97035 Then gError 95496

        'Item não está cadastrado
        If lErro = 97035 Then gError 95497

        'Coloca Descricao do Item de Servico no Grid
        GridItensServico.TextMatrix(GridItensServico.Row, iGrid_Decricao_Col) = objItemServico.sDescricao

        If GridItensServico.Row - GridItensServico.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    ElseIf GridItensServico.Row - GridItensServico.FixedRows < objGridInt.iLinhasExistentes Then
    
       gError 97558

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 95498

    Saida_Celula_MaskItemServico = SUCESSO

    Exit Function

Erro_Saida_Celula_MaskItemServico:

    Saida_Celula_MaskItemServico = gErr

    Select Case gErr

        Case 95495
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_SERVICO_JA_EXISTENTE_GRIDITENSSERVICO", gErr, objItemServico.iCodigo, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 95496, 95498
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 95497
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMSERVICO_NAO_CADASTRADO", gErr, objItemServico.iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 97558
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMSERVICO_NAO_PREENCHIDO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function


Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela

Dim lErro As Long
Dim objServItemServ As New ClassServItemServ

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da coleção para o obj
    objServItemServ.sProduto = colCampoValor.Item("Codigo").vValor

    'Preenche a tela
    lErro = Traz_ServItemServ_Tela(objServItemServ)
    If lErro <> SUCESSO Then gError 95486

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 95486

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub maskServico_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub BotaoItensServico_Click()

Dim objItemServico As New ClassItemServico
Dim sItemServico As String
Dim sItemServico1 As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection

On Error GoTo Erro_BotaoItensServico_Click

    If Me.ActiveControl Is MaskItemServico Then

        sItemServico1 = MaskItemServico.Text

    Else

        'Verifica se tem alguma linha selecionada no Grid
        If GridItensServico.Row = 0 Then gError 95499

        sItemServico1 = GridItensServico.TextMatrix(GridItensServico.Row, iGrid_ItemServico_Col)

    End If

    'preenche o codigo do item de servico
    objItemServico.iCodigo = StrParaInt(sItemServico1)

    'Chama a tela de browse ItensServicoLista
    Call Chama_Tela("ItemServicoLista", colSelecao, objItemServico, objEventoItemServico)

    Exit Sub

Erro_BotaoItensServico_Click:

    Select Case gErr

        Case 95499
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Function Traz_ServItemServ_Tela(objServItemServ As ClassServItemServ)
'Traz os dados para a tela

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objServico As New ClassServico
Dim iIndice As Integer
Dim objItemServico As New ClassItemServico

On Error GoTo Erro_Traz_ServItemServ_Tela

    'Limpa a tela
    Call Limpa_Tela_ServItemServ
    
    'Lê o Produto
    objProduto.sCodigo = objServItemServ.sProduto
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 97502

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 97503
    
    'Coloca na tela o codigo e a descricao do servico
    MaskServico.PromptInclude = False
    MaskServico = objProduto.sCodigo
    MaskServico.PromptInclude = True
    LabelDescricaoServico.Caption = objProduto.sDescricao
    
    'Le todos os itens de servico relacionados com o servico passado como parametro
    objServico.sProduto = objProduto.sCodigo
    
    lErro = CF("ServicoItemServico_Le", objServico)
    If lErro <> SUCESSO And lErro <> 97543 Then gError 97545
    
    'Se nao encontrou => ERRO
    If lErro = 97543 Then gError 97546
    
    'Le todos os itens de servico
    For iIndice = 1 To objServico.colServItemServ.Count
    
        objItemServico.iCodigo = objServico.colServItemServ.Item(iIndice).iCodItemServico
         
        lErro = CF("ItemServico_Le", objItemServico)
        If lErro <> SUCESSO And lErro <> 97035 Then gError 97548

        'Item não está cadastrado
        If lErro = 97035 Then gError 97549
        
        'Coloca no grid os dados do item de servico
        GridItensServico.TextMatrix(iIndice, iGrid_ItemServico_Col) = objItemServico.iCodigo
        GridItensServico.TextMatrix(iIndice, iGrid_Decricao_Col) = objItemServico.sDescricao
        
        objGridItensServico.iLinhasExistentes = iIndice
        
    Next
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Traz_ServItemServ_Tela = SUCESSO

    Exit Function

Erro_Traz_ServItemServ_Tela:

    Traz_ServItemServ_Tela = gErr

    Select Case gErr

        Case 97502, 97545, 97546, 97548

        Case 97503
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 97549
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMSERVICO_NAO_CADASTRADO", gErr, objItemServico.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? criar IDH Parent.HelpContextID = IDH_BAIXA_PARCELAS_RECEBER_TITULOS
    Set Form_Load_Ocx = Me
    Caption = "Serviço X Item de Serviço"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ServItemServ"

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

   RaiseEvent Unload

End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    'm_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******

Function Move_Tela_Memoria(objServico As ClassServico) As Long

Dim iIndice As Integer
Dim objServItemServ As ClassServItemServ
Dim iProdutoPreenchido As Integer
Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_Move_Tela_Memoria

    'Guarda o Código e a descricao do servico no obj
    lErro = CF("Produto_Formata", MaskServico.Text, sProduto, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 97507
    
    objServico.sProduto = sProduto

    'Guarda os itens de servicos associados
    For iIndice = 1 To objGridItensServico.iLinhasExistentes

        'Inicializa o obj
        Set objServItemServ = New ClassServItemServ

        'Guarda o código do item se servico
        objServItemServ.iCodItemServico = CInt(GridItensServico.TextMatrix(iIndice, iGrid_ItemServico_Col))
        objServItemServ.iOrdem = iIndice
        objServItemServ.sProduto = objServico.sProduto

        'Adiciona na Colecao
        objServico.colServItemServ.Add objServItemServ

    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 97507

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Function Form_Activate()

    Call TelaIndice_Preenche(Me)

End Function

Public Function Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objServico As New ClassServico
Dim objServItemServ As New ClassServItemServ

On Error GoTo Erro_Gravar_Registro

    'Coloca o cursor com formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Verificar se os Campos Obrigatórios estão preenchidos
    If Len(Trim(MaskServico.ClipText)) = 0 Then gError 97508
    If objGridItensServico.iLinhasExistentes = 0 Then gError 97509

    'Armazena em objServicos os dados da tela
    lErro = Move_Tela_Memoria(objServico)
    If lErro <> SUCESSO Then gError 97510
    
    objServItemServ.sProduto = objServico.sProduto
    
    lErro = Trata_Alteracao(objServItemServ, objServItemServ.sProduto)
    If lErro <> SUCESSO Then gError 97555

    lErro = CF("ServicoItemServico_Grava", objServico.colServItemServ)
    If lErro <> SUCESSO Then gError 97511

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Retorna o cursor ao formato default
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 97508
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PREENCHIDO", gErr)

        Case 97509
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_LINHA_INEXISTENTE", gErr)

        Case 97510, 97511, 97555

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Retorna o cursor ao formato default
    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function

Private Function Reconfigura_Linha_Grid()

        GridItensServico.BackColorSel = &H8000000D '= Windows HighLight (AZUL)
        GridItensServico.AllowBigSelection = True
        GridItensServico.FocusRect = flexFocusHeavy
        GridItensServico.HighLight = flexHighlightAlways
        GridItensServico.SelectionMode = flexSelectionByRow

End Function

Private Sub LabelServicos_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelServicos_Click

    'Verifica se o serviço foi preenchido
    If Len(Trim(MaskServico.ClipText)) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", MaskServico.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 97553

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ServItemServLista", colSelecao, objProduto, objEventoServicoItemServico)

    Exit Sub

Erro_LabelServicos_Click:

    Select Case gErr

        Case 97553

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub objEventoServicoItemServico_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim bCancel As Boolean
Dim objServItemServ As New ClassServItemServ

On Error GoTo Erro_objEventoServico_evSelecao

    Set objProduto = obj1

    'Faz a leitura dos itens de servico a partir do codigo do servico
    objServItemServ.sProduto = objProduto.sCodigo
    
    lErro = Traz_ServItemServ_Tela(objServItemServ)
    If lErro <> SUCESSO Then gError 97554
    
    'Fecha o Comando de Setas
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Me.Show

    Exit Sub

Erro_objEventoServico_evSelecao:

    Select Case gErr

        Case 97554

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha o Comando de Setas
    Call ComandoSeta_Fechar(Me.Name)

    Exit Sub

End Sub

'Caso o usuario queira acessar o browser através da tecla F3.
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is MaskServico Then
            Call LabelServico_Click
        ElseIf Me.ActiveControl Is MaskItemServico Then
            Call BotaoItensServico_Click
        End If
        
    End If

End Sub
