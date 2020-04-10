VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTTabPrecoItemEmpToda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public gobjInfoUsu As Object

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

'VARIAVEIS GLOBAIS A TELA
Public iAlterado As Integer
Public iProdutoAlterado As Integer

Private WithEvents objEventoItemTabela As AdmEvento
Attribute objEventoItemTabela.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

'--- inicio dos properties get dos controles da tela

Public Property Get Observacao() As Object
     Set Observacao = objUserControl.Controls("Observacao")
End Property

Public Property Get BotaoAtualizaPreco() As Object
     Set BotaoAtualizaPreco = objUserControl.Controls("BotaoAtualizaPreco")
End Property

Public Property Get BotaoFormacaoPreco() As Object
     Set BotaoFormacaoPreco = objUserControl.Controls("BotaoFormacaoPreco")
End Property

Public Property Get TabelaDefault() As Object
     Set TabelaDefault = objUserControl.Controls("TabelaDefault")
End Property

Public Property Get DataVigencia() As Object
     Set DataVigencia = objUserControl.Controls("DataVigencia")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get Tabela() As Object
     Set Tabela = objUserControl.Controls("Tabela")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get BotaoEditarTabela() As Object
     Set BotaoEditarTabela = objUserControl.Controls("BotaoEditarTabela")
End Property

Public Property Get BotaoExcluirTabela() As Object
     Set BotaoExcluirTabela = objUserControl.Controls("BotaoExcluirTabela")
End Property

Public Property Get BotaoCriarTabela() As Object
     Set BotaoCriarTabela = objUserControl.Controls("BotaoCriarTabela")
End Property

Public Property Get Valor() As Object
     Set Valor = objUserControl.Controls("Valor")
End Property

Public Property Get Produto() As Object
     Set Produto = objUserControl.Controls("Produto")
End Property

Public Property Get TabelaLabel() As Object
     Set TabelaLabel = objUserControl.Controls("TabelaLabel")
End Property

Public Property Get DescricaoTabela() As Object
     Set DescricaoTabela = objUserControl.Controls("DescricaoTabela")
End Property

Public Property Get DescricaoProduto() As Object
     Set DescricaoProduto = objUserControl.Controls("DescricaoProduto")
End Property

Public Property Get LabelProduto() As Object
     Set LabelProduto = objUserControl.Controls("LabelProduto")
End Property

Public Property Get UnidadeMedida() As Object
     Set UnidadeMedida = objUserControl.Controls("UnidadeMedida")
End Property

'--- fim dos properties get dos controles da tela

Public Sub BotaoAtualizaPreco_Click()

Dim lErro As Long
Dim objFormacaoPreco As New ClassFormacaoPreco
Dim objFormacaoPrecoTemp As ClassFormacaoPreco
Dim colFormacaoPreco As New Collection
Dim sProduto As String
Dim iPreenchido As Integer, iTabelaPreco As Integer
Dim dValor As Double, objContexto As New ClassContextoPlan

On Error GoTo Erro_BotaoAtualizaPreco_Click

    objFormacaoPreco.iFilialEmpresa = giFilialEmpresa
    objFormacaoPreco.iEscopo = FORMACAO_PRECO_ESCOPO_TABPRECO
    
    If Len(Trim(Tabela.Text)) = 0 Then gError 133035
    If Len(Trim(Produto.ClipText)) = 0 Then gError 133036

    iTabelaPreco = CInt(Tabela.Text)
    objFormacaoPreco.iTabelaPreco = iTabelaPreco

    lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 133037
    
    If iPreenchido = PRODUTO_PREENCHIDO Then objFormacaoPreco.sProduto = sProduto
    
    'Lê a planilha de Formacao de Preço da Tabela/Produto
    lErro = CF("FormacaoPreco_Le1", objFormacaoPreco, colFormacaoPreco)
    If lErro <> SUCESSO And lErro <> 92434 And lErro <> 92432 Then gError 133038
    
    'não há planilha de formação de preço associada
    If lErro <> SUCESSO Then gError 133039

    For Each objFormacaoPrecoTemp In colFormacaoPreco
    
        With objFormacaoPrecoTemp
            .iFilialEmpresa = giFilialEmpresa
            .iEscopo = FORMACAO_PRECO_ESCOPO_TABPRECO
            .iTabelaPreco = iTabelaPreco
            .sProduto = sProduto
        End With
    
    Next
    
    lErro = CF("ContextoPlan_Inicializa", objContexto, objFormacaoPreco.iFilialEmpresa, sProduto, iTabelaPreco, gdtDataAtual, 1)
    If lErro <> SUCESSO Then gError 133040
        
    'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colFormacaoPreco.Item(?).dValor
    lErro = CF("Avalia_Expressao_FPreco1", colFormacaoPreco, dValor, sProduto)
    If lErro <> SUCESSO Then gError 133040

    Valor.Text = Format(dValor, Valor.Format)

    Exit Sub

Erro_BotaoAtualizaPreco_Click:

    Select Case gErr
    
        Case 133035
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)
        
        Case 133036
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
    
        Case 133037, 133038, 133040
        
        Case 133039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMACAOPRECO_NAO_CADASTRADO1", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158308)
    
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoCriarTabela_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objTabelaPreco As New ClassTabelaPreco

    'Chama tela TabelaPrecoCriacao
    Call Chama_Tela_Modal("TabelaPrecoCriacao", objTabelaPreco)
    
    If objTabelaPreco.iCodigo = 0 Then Exit Sub
        
    'Varre a Combo de Tabela de Preço para colocar a Tabela na ordem
    If Tabela.ListCount > 0 Then
    
        For iIndice = 0 To Tabela.ListCount - 1
            If Tabela.ItemData(iIndice) > objTabelaPreco.iCodigo Then
                Exit For
            End If
        Next
    End If
    
    'Preenche a Combo Tabela de Preço com a Tabela Criada
    Tabela.AddItem objTabelaPreco.iCodigo, iIndice
    Tabela.ItemData(iIndice) = objTabelaPreco.iCodigo
    
    Tabela.ListIndex = iIndice
    
    Exit Sub

End Sub

Public Sub BotaoEditarTabela_Click()

Dim lErro As Long
Dim objTabelaPreco As New ClassTabelaPreco

On Error GoTo Error_BotaoEditarTabela_Click

    'Verifica se foi preenchida a ComboBox Tabela
    If Len(Trim(Tabela.Text)) = 0 Then Error 28216

    'Preenche os dados de objTabelaPreco a partir da tela
    objTabelaPreco.iCodigo = CInt(Tabela.Text)
    objTabelaPreco.sDescricao = DescricaoTabela.Caption

    'Chama tela TabelaPrecoAlteracao
    Call Chama_Tela_Modal("TabelaPrecoAlteracao", objTabelaPreco)

    'Coloca Descrição na tela
    DescricaoTabela.Caption = objTabelaPreco.sDescricao

    Exit Sub

Error_BotaoEditarTabela_Click:

    Select Case Err

        Case 28216
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158309)

    End Select

    Exit Sub

End Sub

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim vbMsgRes As VbMsgBoxResult
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento de tabela
    If Len(Trim(Tabela.Text)) = 0 Then Error 28228

    'Verifica preenchimento de produto
    If Len(Trim(Produto.ClipText)) = 0 Then Error 28229

    If Len(Trim(DataVigencia.Text)) = 0 Then Error 20859

    sProduto = Produto.Text

    'Critica o formato do Produto e se existe no BD
    lErro = CF("Produto_Critica2", sProduto, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then Error 28230

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then Error 27204
    
    If lErro = 25041 Then Error 28231

    'Faz FilialEmpresa igual a giFilialEmpresa
    objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa

   'Preenche objTabelaPrecoitem
    objTabelaPrecoItem.iCodTabela = CInt(Tabela.Text)
    objTabelaPrecoItem.sCodProduto = objProduto.sCodigo
    objTabelaPrecoItem.dtDataVigencia = CDate(DataVigencia.Text)
    
    lErro = CF2(Me, "TPI_Move_Tabela_Memoria", objTabelaPrecoItem)
    If lErro <> SUCESSO Then gError 182684

    lErro = CF("TabelaPrecoItem_Le_DataVigente", objTabelaPrecoItem)
    If lErro <> SUCESSO And lErro <> 58503 Then Error 28232

    If lErro = 58503 Then Error 28233

    'Pede confirmação para exclusão de item de tabela de preço
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_ITEM_TABELA_DE_PRECO", objTabelaPrecoItem.sCodProduto, objTabelaPrecoItem.iCodTabela)

    If vbMsgRes = vbYes Then
        
        'Exclui o Item da tabela de Preço
        lErro = CF("TabelaPrecoItem_Exclui", objTabelaPrecoItem)
        If lErro <> SUCESSO Then Error 28235

        'Limpa a tela
        Call Limpa_Tela_TabelaPrecoItemEmpresaToda_Grava

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 20859
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_VIGENCIA_NAO_PREENCHIDA", Err)

        Case 28228
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", Err)

        Case 28229, 27204
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)

        Case 28230, 28232, 28235, 182684

        Case 28233
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_NAO_CADASTRADO", Err, objTabelaPrecoItem.iCodTabela, objTabelaPrecoItem.sCodProduto)

        Case 28231
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158310)

    End Select

    Exit Sub

End Sub

Public Sub BotaoExcluirTabela_Click()

Dim lErro As Long
Dim objTabelaPreco As New ClassTabelaPreco
Dim vbMsgRes As VbMsgBoxResult
Dim colCodigo As New Collection
Dim iIndice As Integer

On Error GoTo Error_BotaoExcluirTabela_Click

    'Verifica se foi preenchida a ComboBox Tabela
    If Len(Trim(Tabela.Text)) = 0 Then Error 28211

    iIndice = Tabela.ListIndex

    objTabelaPreco.iCodigo = CInt(Tabela.Text)
    
    'Verifica se a Tabela existe no BD
    lErro = CF("TabelaPreco_Le", objTabelaPreco)
    If lErro <> SUCESSO And lErro <> 28004 Then Error 28212

    If lErro = 28004 Then Error 28213

    'Pede confirmação da exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_TABELA_DE_PRECO", objTabelaPreco.iCodigo)

    'Se não confirmar, sai
    If vbMsgRes = vbNo Then Exit Sub

    'Exclui a Tabela de Preço
    lErro = CF("TabelaPreco_Exclui", objTabelaPreco)
    If lErro <> SUCESSO Then Error 28214
    
    'Remove da Lista
    Tabela.RemoveItem iIndice
    
    Call Limpa_Tela_TabelaPrecoItemEmpresaToda
    
    Exit Sub

Error_BotaoExcluirTabela_Click:

    Select Case Err

        Case 28211
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", Err)

        Case 28212, 28214, 28215

        Case 28213
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_INEXISTENTE", Err, objTabelaPreco.iCodigo)
            Tabela.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158311)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub BotaoFormacaoPreco_Click()

Dim lErro As Long
Dim objFormacaoPreco As ClassFormacaoPreco
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_BotaoFormacaoPreco_Click

    If Len(Trim(Tabela.Text)) <> 0 And Len(Trim(Produto.ClipText)) <> 0 Then

        Set objFormacaoPreco = New ClassFormacaoPreco

        objFormacaoPreco.iFilialEmpresa = giFilialEmpresa
        objFormacaoPreco.iEscopo = FORMACAO_PRECO_ESCOPO_TABPRECO
    
        objFormacaoPreco.iTabelaPreco = CInt(Tabela.Text)
    
        lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 133034
        
        If iPreenchido = PRODUTO_PREENCHIDO Then objFormacaoPreco.sProduto = sProduto
    
    End If
    
    Call Chama_Tela("FormacaoPreco", objFormacaoPreco)
    
    Exit Sub

Erro_BotaoFormacaoPreco_Click:

    Select Case gErr
    
        Case 133034
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158312)
    
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama Gravar_Registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 28218

    'Limpa Tela
    Call Limpa_Tela_TabelaPrecoItemEmpresaToda_Grava

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 28218
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158313)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 28236

    'Limpa a tela
    Call Limpa_Tela_TabelaPrecoItemEmpresaToda

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 28236

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158314)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim colCodigo As New Collection

On Error GoTo Erro_Form_Load
    
    Set objEventoItemTabela = New AdmEvento
    Set objEventoProduto = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 28183

    'Chama Carrega_TabelaPreco
    lErro = Carrega_TabelaPreco(colCodigo)
    If lErro <> SUCESSO Then gError 28184

''    'Inicializa a Lista de Produtos
''    lErro = CF("Carga_Arvore_Produto_Venda",Produtos.Nodes)
''    If lErro <> SUCESSO Then Error 28185
''
    'Preenche ComboBox Tabela
    If Tabela.ListCount > 0 Then Tabela.ListIndex = 0
    
    Valor.Format = gobjFAT.sFormatoPrecoUnitario
    
    lErro = CF2(Me, "TPI_Form_Load")
    If lErro <> SUCESSO Then gError 182673

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 28183, 28184, 182673

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158315)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
   
    Set objEventoItemTabela = Nothing
    Set objEventoProduto = Nothing
    
    Call CF2(Me, "TPI_Form_UnLoad")
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim iIndice As Integer

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TabelaPrecoItemProduto"

    'Lê os dados da Tela Tabela Preço Item
    lErro = Move_Tela_Memoria(objTabelaPrecoItem)
    If lErro <> SUCESSO Then gError 28186

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "FilialEmpresa", objTabelaPrecoItem.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "CodTabela", objTabelaPrecoItem.iCodTabela, 0, "CodTabela"
    colCampoValor.Add "CodProduto", objTabelaPrecoItem.sCodProduto, STRING_PRODUTO, "CodProduto"
    
    lErro = CF2(Me, "TPI_Tela_Extrai", sTabela, colCampoValor, colSelecao, objTabelaPrecoItem)
    If lErro <> SUCESSO Then gError 182676
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, EMPRESA_TODA

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 28186, 182676

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158316)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_TabelaPrecoItemEmpresaToda()

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Call CF2(Me, "TPI_Limpa_Tela")

    DescricaoProduto.Caption = ""
    Call CF2(Me, "TabPreco_Produto_ExibeInfo", "")
    UnidadeMedida.Caption = ""
    Tabela.ListIndex = -1
    DescricaoTabela.Caption = ""
    TabelaDefault.Value = vbUnchecked
    
    DataVigencia.Clear
    
    'Funcao generica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

End Sub

Sub Limpa_Tela_TabelaPrecoItemEmpresaToda_Grava()

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Call CF2(Me, "TPI_Limpa_Tela")

    DescricaoProduto.Caption = ""
    Call CF2(Me, "TabPreco_Produto_ExibeInfo", "")
    UnidadeMedida.Caption = ""
    DescricaoTabela.Caption = ""

    TabelaDefault.Value = vbUnchecked
    DataVigencia.Clear
    

    'Funcao generica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Produto.SetFocus

End Sub

Private Function Move_Tela_Memoria(objTabelaPrecoItem As ClassTabelaPrecoItem) As Long
'Move os dados da Tela para objTabelaPrecoItem

Dim lErro As Long
Dim objTabelaPreco As New ClassTabelaPreco
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
    
    If Len(Trim(Tabela.Text)) > 0 Then objTabelaPrecoItem.iCodTabela = CInt(Tabela.Text)

    If Len(Trim(Produto.ClipText)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 28192

        objTabelaPrecoItem.sCodProduto = sProdutoFormatado
        
    End If

    'Preenche objTabelaPrecoItem com valor e DataVigencia
    If Len(Trim(Valor.ClipText)) > 0 Then objTabelaPrecoItem.dPreco = CDbl(Valor.ClipText)

    If Len(Trim(DataVigencia.Text)) > 0 Then objTabelaPrecoItem.dtDataVigencia = CDate(DataVigencia.Text)

    If TabelaDefault.Value = vbChecked Then
        objTabelaPrecoItem.iTabelaDefault = Padrao
    Else
        objTabelaPrecoItem.iTabelaDefault = NAO_PADRAO
    End If
    
    objTabelaPrecoItem.sObservacao = Observacao.Text

    Set objTabelaPrecoItem.objTela = Me
    
    lErro = CF2(Me, "TPI_Move_Tabela_Memoria", objTabelaPrecoItem)
    If lErro <> SUCESSO Then gError 182674

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 28192, 182674

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158317)

    End Select

    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem

On Error GoTo Erro_Tela_Preenche

    objTabelaPrecoItem.iCodTabela = colCampoValor.Item("CodTabela").vValor
    objTabelaPrecoItem.sCodProduto = colCampoValor.Item("CodProduto").vValor
    
    lErro = CF2(Me, "TPI_Tela_Preenche", colCampoValor, objTabelaPrecoItem)
    If lErro <> SUCESSO Then gError 182677

    If (objTabelaPrecoItem.iCodTabela <> 0) And (objTabelaPrecoItem.sCodProduto <> "") Then

        'Carrega objTabelaPrecoItem com os dados passados em colCampoValor
        objTabelaPrecoItem.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor

        'Traz os dados para tela
        lErro = Traz_TabelaPrecoItemEmpresaToda_Tela(objTabelaPrecoItem)
        If lErro <> SUCESSO Then gError 28195

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 28195, 182677

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158318)

    End Select

    Exit Sub

End Sub

Private Function Traz_TabelaPrecoItemEmpresaToda_Tela(objTabelaPrecoItem As ClassTabelaPrecoItem) As Long
'Traz os dados da Tabela Preço Item Empresa Toda para a Tela

Dim lErro As Long
Dim sCodigo As String
Dim sProdutoEnxuto As String
Dim objTabelaPreco As New ClassTabelaPreco

On Error GoTo Erro_Traz_TabelaPrecoItemEmpresaToda_Tela
    
    'Limpa a tela
    Call Limpa_Tela_TabelaPrecoItemEmpresaToda
    
    'Seleciona a Tabela na combo
    Call Combo_Seleciona_ItemData(Tabela, objTabelaPrecoItem.iCodTabela)
    
    sCodigo = objTabelaPrecoItem.sCodProduto
    
    'Preenche o Produto
    lErro = CF("Traz_Produto_MaskEd", sCodigo, Produto, DescricaoProduto)
    If lErro <> SUCESSO Then gError 28196
    
    lErro = CF2(Me, "TPI_Traz_Tabela_Tela", objTabelaPrecoItem)
    If lErro <> SUCESSO Then gError 182685
    
    'Chama o Validate do Produto
    Call Produto_Validate(bSGECancelDummy)

    Observacao.Text = objTabelaPrecoItem.sObservacao
    
    iAlterado = 0

    Traz_TabelaPrecoItemEmpresaToda_Tela = SUCESSO

    Exit Function

Erro_Traz_TabelaPrecoItemEmpresaToda_Tela:

    Traz_TabelaPrecoItemEmpresaToda_Tela = gErr

    Select Case gErr

        Case 28196, 182685

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158319)

    End Select

    Exit Function

End Function

Private Function Carrega_TabelaPreco(colCodigo As Collection) As Long
'Carrega a ComboBox Tabela

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Carrega_TabelaPreco

    'Preenche a ComboBox com  os Tipos de Documentos existentes no BD
    lErro = CF("TabelasPreco_Le_Codigos", colCodigo)
    If lErro <> SUCESSO Then Error 28198

    If colCodigo.Count > 0 Then

        For iIndice = 1 To colCodigo.Count

            'Preenche a ComboBox Tabela com os objetos da colecao colTabelaPreco
            Tabela.AddItem colCodigo(iIndice)
            Tabela.ItemData(Tabela.NewIndex) = colCodigo(iIndice)

        Next

    End If

    Carrega_TabelaPreco = SUCESSO

    Exit Function

Erro_Carrega_TabelaPreco:

    Carrega_TabelaPreco = Err

    Select Case Err

        Case 28198

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158320)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objTabelaPrecoItem As ClassTabelaPrecoItem) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um item selecionado, exibir seus dados
    If Not (objTabelaPrecoItem Is Nothing) Then

        'Alimenta FilialEmpresa no obj com EMPRESA_TODA
        objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA

        'Verifica se ítem existe
        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then Error 28199

        If lErro = 28014 Then Error 28200
    
        'Preenche a Tela com o item da Tabela de Preço
        lErro = Traz_TabelaPrecoItemEmpresaToda_Tela(objTabelaPrecoItem)
        If lErro <> SUCESSO Then Error 28201

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 28199, 28201

        Case 28200
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_NAO_CADASTRADO", Err, objTabelaPrecoItem.iCodTabela, objTabelaPrecoItem.sCodProduto)
            Call Limpa_Tela_TabelaPrecoItemEmpresaToda

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158321)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub LabelProduto_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sBrowse As String
Dim objProduto As New ClassProduto, objTabelaPreco As New ClassTabelaPreco
Dim colSelecao As New Collection, bCompra As Boolean

On Error GoTo Erro_LabelProduto_Click

    'Verifica se o produto foi preenchido
    If Len(Produto.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82707

        objProduto.sCodigo = sProdutoFormatado

    End If
    
    bCompra = False
    
    'Verifica se Tabela está preenchida
    If Len(Trim(Tabela.Text)) > 0 Then
    
        objTabelaPreco.iCodigo = CInt(Tabela.Text)
        lErro = CF("TabelaPreco_le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 82673
        If lErro = SUCESSO Then
            If objTabelaPreco.iTipo = TABELA_PRECO_TIPO_COMPRA Then bCompra = True
        End If
    
    End If
    
    If bCompra Then
    
        sBrowse = "ProdutoCompraLista"
    
    Else
    
        sBrowse = "ProdutoVendaTodosLista"
    
    End If
    
    Call CF2(Me, "TPI_BotaoProdutos_Click", sBrowse)

    Call Chama_Tela(sBrowse, colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_LabelProduto_Click:

    Select Case gErr

        Case 82707

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158322)

    End Select

    Exit Sub

End Sub

Private Sub objEventoItemTabela_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTabelaPrecoItem As ClassTabelaPrecoItem

On Error GoTo Erro_objEventoItemTabela_evSelecao

    Set objTabelaPrecoItem = obj1
    
    Call CF2(Me, "TPI_Trata_Retorno_Browse", objTabelaPrecoItem)
    
    'Preenche os campos da Tela com o item da Tabela de Preço
    lErro = Traz_TabelaPrecoItemEmpresaToda_Tela(objTabelaPrecoItem)
    If lErro <> SUCESSO Then Error 49867
        
    Me.Show

    Exit Sub

Erro_objEventoItemTabela_evSelecao:

    Select Case Err

        Case 49867
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158323)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82708

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82709
    
    Produto.PromptInclude = False
    Produto.Text = objProduto.sCodigo
    Produto.PromptInclude = True
    
    DescricaoProduto.Caption = objProduto.sDescricao
    Call CF2(Me, "TabPreco_Produto_ExibeInfo", objProduto.sCodigo, objProduto.sCor)
    
    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 82708

        Case 82709
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158324)

    End Select

    Exit Sub

End Sub

Public Sub Produto_Change()

    iProdutoAlterado = 1
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String
Dim iProdutoPreenchido As Integer, objTabelaPreco As New ClassTabelaPreco

On Error GoTo Erro_Produto_Validate

    If iProdutoAlterado = 1 Then
        
        If Len(Trim(Produto.ClipText)) > 0 Then
    
            sProduto = Produto.Text
    
            'Critica o formato do Produto
            lErro = CF("Produto_Critica2", sProduto, objProduto, iProdutoPreenchido)
            If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 28219
    
            If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 27206
        
            If lErro = 25041 Then gError 28220
            
            'Preenche ProdutoDescricao com Descrição do Produto
            DescricaoProduto.Caption = objProduto.sDescricao
            Call CF2(Me, "TabPreco_Produto_ExibeInfo", objProduto.sCodigo, objProduto.sCor)
            
            'Limpa a Combo de Data
            DataVigencia.Clear
            
            UnidadeMedida.Caption = ""
            
            'Verifica se Tabela está preenchida
            If Len(Trim(Tabela.Text)) > 0 Then
                
                'Carrega a Combo de Data de Vigência
                lErro = Carrega_Combo_DataVigencia(objProduto)
                If lErro <> SUCESSO And lErro <> 58521 Then gError 58511
                
                'Preenche objTabelaPrecoItem
                objTabelaPrecoItem.iCodTabela = CInt(Tabela.Text)
                objTabelaPrecoItem.sCodProduto = objProduto.sCodigo
                objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
                
                If objTabelaPrecoItem.iCodTabela = objProduto.iTabelaPreco Then
                    TabelaDefault.Value = vbChecked
                Else
                    TabelaDefault.Value = vbUnchecked
                End If
                
                lErro = CF2(Me, "TPI_Move_Tabela_Memoria", objTabelaPrecoItem)
                If lErro <> SUCESSO Then gError 182680
    
                'Lê TabelaPrecoItem
                lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
                If lErro <> SUCESSO And lErro <> 28014 Then gError 28243
                
                'Se encontrou o Ítem mostra o Valor na tela
                If lErro <> 28014 Then
                    
                    'Preenche o Valor e Seleciona a Data de Vigencia
                    Valor.Text = Format(objTabelaPrecoItem.dPreco, Valor.Format)
                    If objTabelaPrecoItem.dtDataVigencia <> DATA_NULA Then
                        DataVigencia.Text = Format(objTabelaPrecoItem.dtDataVigencia, "dd/mm/yyyy")
                        Call Combo_Item_Igual(DataVigencia)
                    End If
                                    
                Else
                
                    'Se não encontrou o Ítem limpa o campo Valor e Coloca a Data Atual em Data de Vigência
                    Valor.Text = ""
                    DataVigencia.Text = Format(gdtDataAtual, "dd/mm/yyyy")
                
                End If
            
                objTabelaPreco.iCodigo = CInt(Tabela.Text)
                lErro = CF("TabelaPreco_le", objTabelaPreco)
                If lErro <> SUCESSO And lErro <> 28004 Then gError 182680
                If lErro = SUCESSO Then
                    If objTabelaPreco.iTipo = TABELA_PRECO_TIPO_COMPRA Then UnidadeMedida.Caption = objProduto.sSiglaUMCompra
                End If

            End If
    
            'Mostra a unidade de medida na tela
            If UnidadeMedida.Caption = "" Then
                
                If objProduto.iFaturamento = PRODUTO_NAO_VENDAVEL Then gError 62094
                
                'Mostra a unidade de medida na tela
                UnidadeMedida.Caption = objProduto.sSiglaUMVenda
            
            End If
            
        Else
            'Limpa DescricaoProduto
            DescricaoProduto.Caption = ""
            Call CF2(Me, "TabPreco_Produto_ExibeInfo", "")
            
            'Limpa unidade de medida do produto
            UnidadeMedida.Caption = ""

        End If
        
        lErro = CF2(Me, "TPI_Traz_Tabela_Tela", objTabelaPrecoItem)
        If lErro <> SUCESSO Then gError 182675
        
        iProdutoAlterado = 0
        
    End If
    
    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 28219, 27206, 182675, 182680

        Case 28220 'Não encontrou Produto no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Produto", objProduto)

            Else
                'Limpa DescricaoProduto
                DescricaoProduto.Caption = ""
                Call CF2(Me, "TabPreco_Produto_ExibeInfo", "")
                'Segura o foco
            End If
        
        Case 28243, 58511
        
        Case 62094
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_VENDIDO", gErr, Produto.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158325)

    End Select

    Exit Sub

End Sub

Public Sub Tabela_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Tabela_Click()

Dim lErro As Long
Dim objTabelaPreco As New ClassTabelaPreco
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Error_Tabela_Click

    'Verifica se foi preenchida a ComboBox Tabela
    If Tabela.ListIndex <> -1 Then

        objTabelaPreco.iCodigo = CInt(Tabela.Text)
        
        'Lê a Tabela no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then Error 28202
        
        'Se não encontrar ----> Erro
        If lErro = 28004 Then Error 28203
        
        'Preenche a Descrição
        DescricaoTabela.Caption = objTabelaPreco.sDescricao
    
        If Len(Trim(Produto.ClipText)) > 0 Then
        
            sProduto = Produto.Text
        
            'Critica o Produto
            lErro = CF("Produto_Critica2", sProduto, objProduto, iProdutoPreenchido)
            If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then Error 28241
            
            If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then Error 27207
            
            
            'Carrega a Combo de Data de Vigência
            lErro = Carrega_Combo_DataVigencia(objProduto)
            If lErro <> SUCESSO And lErro <> 58521 Then Error 58512
            
            'Preenche objTabelaPrecoItem
            objTabelaPrecoItem.sCodProduto = objProduto.sCodigo
            objTabelaPrecoItem.iCodTabela = CInt(Tabela.Text)
            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
            
            If objProduto.iTabelaPreco = objTabelaPrecoItem.iCodTabela Then
                TabelaDefault.Value = vbChecked
            Else
                TabelaDefault.Value = vbUnchecked
            End If
            
            'Lê TabelaPrecoItem
            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then Error 28242
            
            'Se encontrou o Ítem, mostra Valor na tela
            If lErro <> 28014 Then
                Valor.Text = Format(objTabelaPrecoItem.dPreco, Valor.Format)
                
                'Preenche a Data de Vigência
                If objTabelaPrecoItem.dtDataVigencia <> DATA_NULA Then
                    DataVigencia.Text = Format(objTabelaPrecoItem.dtDataVigencia, "dd/mm/yyyy")
                    Call Combo_Item_Igual(DataVigencia)
                End If

            Else
                
                'Se não encontrou, limpa o campo Valor e coloca Data Atual em Data de Vigência
                Valor.Text = ""
                DataVigencia.Text = Format(gdtDataAtual, "dd/mm/yyyy")

                
            End If
            
        End If
    
    Else
        
        DescricaoTabela.Caption = ""

    End If

    iAlterado = 0

    Exit Sub

Error_Tabela_Click:

    Select Case Err

        Case 28202, 28241, 28242, 27207, 58512

        Case 28203
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_INEXISTENTE", Err, objTabelaPreco.iCodigo)
            Tabela.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158326)

    End Select

    Exit Sub

End Sub

Public Sub TabelaLabel_Click()

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim colSelecao As New Collection
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_TabelaLabel_Click
    
    If Trim(Produto.ClipText) <> "" Then
        
        'Passa para o formato do BD
        lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then Error 49866

        'Testa se o codigo está preenchido
        If iPreenchido = PRODUTO_PREENCHIDO Then
            objTabelaPrecoItem.sCodProduto = sProduto
        Else
            objTabelaPrecoItem.sCodProduto = ""
        End If
        
    Else
        objTabelaPrecoItem.sCodProduto = ""
    End If
    
    colSelecao.Add StrParaInt(Tabela.Text)

    Call Chama_Tela("TabPrecoItemLista", colSelecao, objTabelaPrecoItem, objEventoItemTabela)

    Exit Sub

Erro_TabelaLabel_Click:

    Select Case Err

        Case 49866

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 158327)

    End Select

    Exit Sub

End Sub

Public Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(Valor.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_Positivo_Critica(Valor.Text)
    If lErro <> SUCESSO Then Error 28204

    Valor.Text = Format(Valor.Text, Valor.Format)

    Exit Sub

Erro_Valor_Validate:

    Cancel = True


    Select Case Err

        Case 28204

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158328)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
'Dim objProduto As New ClassProduto
'Dim vbMsgRes As VbMsgBoxResult
'Dim sProduto As String
'Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento de tabela
    If Len(Trim(Tabela.Text)) = 0 Then gError 28221

    'Verifica preenchimento de produto
    If Len(Trim(Produto.ClipText)) = 0 Then gError 28222

    'Verifica preenchimento de valor
    If Len(Trim(Valor.Text)) = 0 Then gError 28226

    'Verifica o Preenchimento da Data de Vigencia
    If Len(Trim(DataVigencia.Text)) = 0 Then gError 58525
    
    'Verifica se a data de Vigencia é maior que date
    If CDate(DataVigencia) < Date Then gError 58526

    'Chama Move_Tela_Memoria
    lErro = Move_Tela_Memoria(objTabelaPrecoItem)
    If lErro <> SUCESSO Then gError 28227
    
    lErro = CF2(Me, "TPI_Grava_Registro", objTabelaPrecoItem)
    If lErro <> SUCESSO Then gError 182678

'    lErro = Trata_Alteracao(objTabelaPrecoItem, objTabelaPrecoItem.iCodTabela, objTabelaPrecoItem.iFilialEmpresa, objTabelaPrecoItem.sCodProduto, objTabelaPrecoItem.dtDataVigencia)
'    If lErro <> SUCESSO Then gError 32331

'    sProduto = Produto.Text
'
'    'Critica o formato do Produto
'    lErro = CF("Produto_Critica2", sProduto, objProduto, iProdutoPreenchido)
'    If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 126834
'
'    If lErro = 25043 And objProduto.iKitVendaComp <> MARCADO Then
'        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_PRODUTO_GERENCIAL_TABPRECO", Produto.Text)
'
'        If vbMsgRes = vbNo Then gError 126835
'
'    End If

    'Chama TabelaPrecoItem_Grava
    lErro = CF("TabelaPrecoItem_Grava", objTabelaPrecoItem)
    If lErro <> SUCESSO Then gError 28228

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 28221
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)

        Case 28222
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 28226
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO1", gErr)

        Case 28227, 28228, 32331, 182678 ', 126834, 126835

        Case 58525
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_VIGENCIA_NAO_PREENCHIDA", gErr)
        
        Case 58526
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_VIGENCIA_MENOR_DATA_ATUAL", gErr, Date)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158329)

    End Select

    Exit Function

End Function


' - - - - - - Raphael 15/10/99
Public Sub DataVigencia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataVigencia_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataVigencia_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataVigencia.Text)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataVigencia.Text)
    If lErro <> SUCESSO Then Error 58513
    
    Call DataVigencia_Click
    
    Exit Sub

Erro_DataVigencia_Validate:

    Cancel = True


    Select Case Err

        'se houve erro de crítica, segura o foco
        Case 58513

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158330)

    End Select

    Exit Sub

End Sub

Public Sub DataVigencia_Click()

Dim lErro As Long
Dim sProduto As String
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_DataVigencia_Click

    'Verifica se foi preenchida a ComboBox Data de Vigencia
    If DataVigencia.ListIndex >= 0 And Len(Trim(DataVigencia.Text)) > 0 Then

        If Len(Trim(Produto.ClipText)) > 0 Then
        
            sProduto = Produto.Text
        
            'Critica o Produto
            lErro = CF("Produto_Critica2", sProduto, objProduto, iProdutoPreenchido)
            If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then Error 58514
            
            If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then Error 58515
                
            'Se a Tabela de
            If Tabela.ListIndex < 0 Then Error 58516
                
            'Preenche os dados do item da Tabela na Tela com a Data de Vigência escolhida
            lErro = Carrega_Dados_Tabela_DataVigencia(objProduto)
            If lErro <> SUCESSO Then Error 58517
                
        End If
    
    End If

    Exit Sub
    
Erro_DataVigencia_Click:

    Select Case Err
        
        Case 58515, 58516 'Produto ou Tabela não Preenchidos
            
        Case 58514, 58517 'Tratados nas rotinas chamadas
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158331)

    End Select

    Exit Sub
        
End Sub

Function Carrega_Dados_Tabela_DataVigencia(objProduto As ClassProduto) As Long
'Carrega a Tela com os dados do item com a Data de Vigencia Selecionada

Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim lErro As Long

On Error GoTo Erro_Carrega_Dados_Tabela_DataVigencia

    objTabelaPrecoItem.iCodTabela = CInt(Tabela.List(Tabela.ListIndex))
    objTabelaPrecoItem.dtDataVigencia = CDate(DataVigencia.Text)
    objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
    objTabelaPrecoItem.sCodProduto = objProduto.sCodigo
    
    lErro = CF2(Me, "TPI_Move_Tabela_Memoria", objTabelaPrecoItem)
    If lErro <> SUCESSO Then gError 182682
    
    'Lê o Valor no Bd
    lErro = CF("TabelaPrecoItem_Le_DataVigente", objTabelaPrecoItem)
    If lErro <> SUCESSO And lErro <> 58503 Then gError 58518
    
    If lErro = 58503 Then gError 58519
    
    If objTabelaPrecoItem.dPreco <> 0 Then Valor.Text = Format(objTabelaPrecoItem.dPreco, Valor.Format)
        
    Observacao.Text = objTabelaPrecoItem.sObservacao
    
    iAlterado = 0
    
    Carrega_Dados_Tabela_DataVigencia = SUCESSO
    
    Exit Function
    
Erro_Carrega_Dados_Tabela_DataVigencia:

    Carrega_Dados_Tabela_DataVigencia = gErr
    
    Select Case gErr
    
        Case 58518, 182682
        
        Case 58519
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECOITEM_INEXISTENTE2", gErr, objTabelaPrecoItem.iCodTabela, objTabelaPrecoItem.sCodProduto, objTabelaPrecoItem.dtDataVigencia)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158332)

    End Select

    Exit Function
    
End Function

Function Carrega_Combo_DataVigencia(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim colDatasVigentes As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Combo_DataVigencia
    
    objTabelaPrecoItem.sCodProduto = objProduto.sCodigo
    objTabelaPrecoItem.iCodTabela = CInt(Tabela.List(Tabela.ListIndex))
    objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF2(Me, "TPI_Move_Tabela_Memoria", objTabelaPrecoItem)
    If lErro <> SUCESSO Then gError 58520
    
    'Lê as Data de Vigência para o par (Produto e Tabela)
    lErro = CF("TabelaDePrecoItens_Le_DatasVigentes", objTabelaPrecoItem, colDatasVigentes)
    If lErro <> SUCESSO And lErro <> 58483 Then Error 58520
    
    'Se nao encontrou ---> sai
    If lErro = 58483 Then Error 58521
    
    'Limpa a combo de Data
    DataVigencia.Clear
    
    'Preenche a combo de Data
    For iIndice = 1 To colDatasVigentes.Count
        
        DataVigencia.AddItem Format(colDatasVigentes.Item(iIndice), "dd/mm/yyyy")
    
    Next

    Carrega_Combo_DataVigencia = SUCESSO
    
    Exit Function
    
Erro_Carrega_Combo_DataVigencia:
    
    Carrega_Combo_DataVigencia = Err
    
    Select Case Err
        
        Case 58520
        
        Case 58521
            Valor.Text = ""
            DataVigencia.Text = Format(gdtDataAtual, "dd/mm/yyyy")
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158333)

    End Select

    Exit Function
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ITENS_TABELA_PRECO_EMPRESA_TODA
    Set Form_Load_Ocx = Me
    Caption = "Itens da Tabela de Preço - Empresa Toda"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TabelaPrecoItemEmpresaToda"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
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

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Tabela Then
            Call TabelaLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call LabelProduto_Click
        End If
    
    End If

End Sub

Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

'Devolve Parent do User Control
Public Property Get Parent() As Object
    Set Parent = objUserControl.Parent
End Property

Public Property Get Controls() As Object
    Set Controls = objUserControl.Controls
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = objUserControl.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = objUserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    objUserControl.Enabled = New_Enabled
End Property

'#################################################
'Inserido por Wagner 18/05/2006
Public Sub BotaoKitVenda_Click()

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoKitVenda_Click

    sProduto1 = Produto.Text
   
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 177580
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto

    'Chama a Tela ProdutoVendaLista
    Call Chama_Tela("ProdutoKitVendaLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub
    
Erro_BotaoKitVenda_Click:

    Select Case gErr
    
        Case 177579
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 177580
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177581)
    
    End Select
    
    Exit Sub
    
End Sub
'##############################################################

Public Function TPI_Form_Load(ByVal objCT As Object) As Long
    TPI_Form_Load = SUCESSO
End Function

Function TPI_Move_Tabela_Memoria(ByVal objCT As Object, objTabelaPrecoItem As ClassTabelaPrecoItem) As Long
    TPI_Move_Tabela_Memoria = SUCESSO
End Function

Function TPI_Traz_Tabela_Tela(ByVal objCT As Object, objTabelaPrecoItem As ClassTabelaPrecoItem) As Long
    TPI_Traz_Tabela_Tela = SUCESSO
End Function

Function TPI_Tela_Preenche(ByVal objCT As Object, colCampoValor As AdmColCampoValor, ByVal objTabelaPrecoItem As ClassTabelaPrecoItem) As Long
    TPI_Tela_Preenche = SUCESSO
End Function

Function TPI_Tela_Extrai(ByVal objCT As Object, sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro, ByVal objTabelaPrecoItem As ClassTabelaPrecoItem) As Long
    TPI_Tela_Extrai = SUCESSO
End Function

Function TPI_Grava_Registro(ByVal objCT As Object, ByVal objTabelaPrecoItem As ClassTabelaPrecoItem) As Long

Dim lErro As Long
Dim sProduto As String
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TPI_Grava_Registro

    lErro = Trata_Alteracao(objTabelaPrecoItem, objTabelaPrecoItem.iCodTabela, objTabelaPrecoItem.iFilialEmpresa, objTabelaPrecoItem.sCodProduto, objTabelaPrecoItem.dtDataVigencia)
    If lErro <> SUCESSO Then gError 32331

    sProduto = Produto.Text

    'Critica o formato do Produto
    lErro = CF("Produto_Critica2", sProduto, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 126834

    If lErro = 25043 And objProduto.iKitVendaComp <> MARCADO And Len(Trim(objProduto.sGrade)) = 0 Then
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_PRODUTO_GERENCIAL_TABPRECO", Produto.Text)

        If vbMsgRes = vbNo Then gError 126835

    End If
    
    TPI_Grava_Registro = SUCESSO
    
    Exit Function

Erro_TPI_Grava_Registro:

    TPI_Grava_Registro = gErr

    Select Case gErr
    
        Case 32331, 126834, 126835
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182669)

    End Select

    Exit Function
    
End Function

Function TPI_Limpa_Tela(ByVal objCT As Object) As Long
    TPI_Limpa_Tela = SUCESSO
End Function

Public Function TPI_Form_UnLoad(ByVal objCT As Object) As Long
    TPI_Form_UnLoad = SUCESSO
End Function

Function TPI_Trata_Retorno_Browse(ByVal objCT As Object, ByVal objTabelaPrecoItem As ClassTabelaPrecoItem) As Long
    TPI_Trata_Retorno_Browse = SUCESSO
End Function

Public Function TPI_BotaoProdutos_Click(ByVal objCT As Object, sNomeBrowse As String) As Long
    TPI_BotaoProdutos_Click = SUCESSO
End Function

Public Sub Observacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

