VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTTabelaPrecoItem"
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

Dim iAlterado As Integer
Dim iProdutoAlterado As Integer
Dim dtDataVigenciaAnt As Date

Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoItemTabela As AdmEvento
Attribute objEventoItemTabela.VB_VarHelpID = -1

'workflow
Private Const ALTTABPRECO As String = "AltTabPreco"
Private Const ALTTABPRECODESC As String = "AltTabPrecoDesc"


'--- inicio dos properties get dos controles da tela

Public Property Get Observacao() As Object
     Set Observacao = objUserControl.Controls("Observacao")
End Property

Public Property Get BotaoFormacaoPreco() As Object
     Set BotaoFormacaoPreco = objUserControl.Controls("BotaoFormacaoPreco")
End Property

Public Property Get BotaoAtualizaPreco() As Object
     Set BotaoAtualizaPreco = objUserControl.Controls("BotaoAtualizaPreco")
End Property

Public Property Get TabelaDefault() As Object
     Set TabelaDefault = objUserControl.Controls("TabelaDefault")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get BotaoImportarItens() As Object
     Set BotaoImportarItens = objUserControl.Controls("BotaoImportarItens")
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

Public Property Get Valor() As Object
     Set Valor = objUserControl.Controls("Valor")
End Property

Public Property Get Produto() As Object
     Set Produto = objUserControl.Controls("Produto")
End Property

Public Property Get ProdutoLabel() As Object
     Set ProdutoLabel = objUserControl.Controls("ProdutoLabel")
End Property

Public Property Get DescricaoProduto() As Object
     Set DescricaoProduto = objUserControl.Controls("DescricaoProduto")
End Property

Public Property Get DescricaoTabela() As Object
     Set DescricaoTabela = objUserControl.Controls("DescricaoTabela")
End Property

Public Property Get TabelaLabel() As Object
     Set TabelaLabel = objUserControl.Controls("TabelaLabel")
End Property

Public Property Get ValorEmpresa() As Object
     Set ValorEmpresa = objUserControl.Controls("ValorEmpresa")
End Property

Public Property Get UnidadeMedida() As Object
     Set UnidadeMedida = objUserControl.Controls("UnidadeMedida")
End Property

Public Property Get PrecoComDesconto() As Object
     Set PrecoComDesconto = objUserControl.Controls("PrecoComDesconto")
End Property

Public Property Get PercDesconto() As Object
     Set PercDesconto = objUserControl.Controls("PercDesconto")
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
    
    If Len(Trim(Tabela.Text)) = 0 Then gError 92424
    If Len(Trim(Produto.ClipText)) = 0 Then gError 92425

    iTabelaPreco = CInt(Tabela.Text)
    objFormacaoPreco.iTabelaPreco = iTabelaPreco

    lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 92426
    
    If iPreenchido = PRODUTO_PREENCHIDO Then objFormacaoPreco.sProduto = sProduto
    
    'Lê a planilha de Formacao de Preço da Tabela/Produto
    lErro = CF("FormacaoPreco_Le1", objFormacaoPreco, colFormacaoPreco)
    If lErro <> SUCESSO And lErro <> 92434 And lErro <> 92432 Then gError 92427
    
    'não há planilha de formação de preço associada
    If lErro <> SUCESSO Then gError 92478

    For Each objFormacaoPrecoTemp In colFormacaoPreco
    
        With objFormacaoPrecoTemp
            .iFilialEmpresa = giFilialEmpresa
            .iEscopo = FORMACAO_PRECO_ESCOPO_TABPRECO
            .iTabelaPreco = iTabelaPreco
            .sProduto = sProduto
        End With
    
    Next
    
    lErro = CF("ContextoPlan_Inicializa", objContexto, objFormacaoPreco.iFilialEmpresa, sProduto, iTabelaPreco, gdtDataAtual, 1)
    If lErro <> SUCESSO Then gError 92428
        
    'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colFormacaoPreco.Item(?).dValor
    lErro = CF("Avalia_Expressao_FPreco1", colFormacaoPreco, dValor, sProduto, objContexto)
    If lErro <> SUCESSO Then gError 92428

    Valor.Text = Format(dValor, Valor.Format)
    
    Call Trata_Desconto

    Exit Sub

Erro_BotaoAtualizaPreco_Click:

    Select Case gErr
    
        Case 92424
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)
        
        Case 92425
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
    
        Case 92426, 92427, 92428
        
        Case 92478
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMACAOPRECO_NAO_CADASTRADO1", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158280)
    
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoCriarTabela_Click()

Dim objTabelaPreco As New ClassTabelaPreco
Dim iIndice As Integer

    'Chama tela TabelaPrecoCriacao
    Call Chama_Tela_Modal("TabelaPrecoCriacao", objTabelaPreco)
    
    'Se não criou a Tabela -- > Sai
    If objTabelaPreco.iCodigo = 0 Then Exit Sub
        
    'Procura na Combo de Tabela o indice que a Tabela vai entrar
    If Tabela.ListCount > 0 Then
    
        For iIndice = 0 To Tabela.ListCount - 1
            If Tabela.ItemData(iIndice) > objTabelaPreco.iCodigo Then
                Exit For
            End If
        Next
    End If
    
    'Adiciona na Combo a Tabela Criada
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
    If Len(Trim(Tabela.Text)) = 0 Then Error 42744

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

        Case 42744
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158281)

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
    If Len(Trim(Tabela.Text)) = 0 Then Error 28069

    'Verifica preenchimento de produto
    If Len(Trim(Produto.ClipText)) = 0 Then Error 28070
    
    If Len(Trim(DataVigencia.Text)) = 0 Then Error 58507
    
    sProduto = Produto.Text

    'Critica o formato do Produto e se existe no BD
    lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 Then Error 28156

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then Error 27203
    
    If lErro = 25041 Then Error 28158

    'Faz FilialEmpresa igual a giFilialEmpresa
    objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa

   'Preenche objTabelaPrecoitem
    objTabelaPrecoItem.iCodTabela = CInt(Tabela.Text)
    objTabelaPrecoItem.sCodProduto = objProduto.sCodigo
    objTabelaPrecoItem.dtDataVigencia = CDate(DataVigencia)
    
    lErro = CF("TabelaPrecoItem_Le_DataVigente", objTabelaPrecoItem)
    If lErro <> SUCESSO And lErro <> 58503 Then Error 28071

    If lErro = 58503 Then Error 28072

    'Pede confirmação para exclusão de item de tabela de preço
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_ITEM_TABELA_DE_PRECO", objTabelaPrecoItem.sCodProduto, objTabelaPrecoItem.iCodTabela)

    If vbMsgRes = vbYes Then

        lErro = CF("TabelaPrecoItem_Exclui", objTabelaPrecoItem)
        If lErro <> SUCESSO Then Error 28234

        'Limpa a tela
        Call Limpa_Tela_TabelaPrecoItem_Grava

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 28069
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", Err)

        Case 28070, 27203
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)

        Case 28071, 28156, 28234

        Case 28072
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_NAO_CADASTRADO", Err, objTabelaPrecoItem.iCodTabela, objTabelaPrecoItem.sCodProduto)

        Case 28158
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)
    
        Case 58507
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_VIGENCIA_NAO_PREENCHIDA", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158282)

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
    If Len(Trim(Tabela.Text)) = 0 Then Error 42745

    iIndice = Tabela.ListIndex
    
    objTabelaPreco.iCodigo = CInt(Tabela.Text)

    lErro = CF("TabelaPreco_Le", objTabelaPreco)
    If lErro <> SUCESSO And lErro <> 28004 Then Error 42746

    If lErro = 28004 Then Error 42747

    'Pede confirmação da exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_TABELA_DE_PRECO", objTabelaPreco.iCodigo)

    'Se não confirmar, sai
    If vbMsgRes = vbNo Then Exit Sub

    'Exclui a Tabela de Preço
    lErro = CF("TabelaPreco_Exclui", objTabelaPreco)
    If lErro <> SUCESSO Then Error 42748

    Tabela.RemoveItem iIndice
    
    Call Limpa_Tela_TabelaPrecoItem
    
    Exit Sub

Error_BotaoExcluirTabela_Click:

    Select Case Err

        Case 42745
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", Err)

        Case 42746, 42748

        Case 42747
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_INEXISTENTE", Err, objTabelaPreco.iCodigo)
            Tabela.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158283)

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
        If lErro <> SUCESSO Then gError 92429
        
        If iPreenchido = PRODUTO_PREENCHIDO Then objFormacaoPreco.sProduto = sProduto
    
    End If
    
    Call Chama_Tela("FormacaoPreco", objFormacaoPreco)
    
    Exit Sub

Erro_BotaoFormacaoPreco_Click:

    Select Case gErr
    
        Case 92429
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158284)
    
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 28000

    Call Limpa_Tela_TabelaPrecoItem_Grava

    Produto.SetFocus

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 28000
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158285)

    End Select

    Exit Sub

End Sub

'###################################################################
Public Sub BotaoImportarItens_Click()
'Importa um conjunto de produtos/preços para uma determinada tabela de preços
'e uma data de vigência
'Inserido por Wagner 19/10/2004

Dim lErro As Long
'Dim colTabelaPrecoItem As New Collection
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem

On Error GoTo Erro_BotaoImportarItens_Click

'    GL_objMDIForm.MousePointer = vbHourglass
'
    'Verifica preenchimento de tabela
    If Len(Trim(Tabela.Text)) = 0 Then gError 131015

    'Verifica o Preenchimento da Data de Vigencia
    If Len(Trim(DataVigencia.Text)) = 0 Then gError 131016

    'Verifica se a data de Vigencia é maior que date
    If CDate(DataVigencia) < Date Then gError 131017
    
    objTabelaPrecoItem.dtDataVigencia = StrParaDate(DataVigencia.Text)
    objTabelaPrecoItem.iCodTabela = CInt(Tabela.Text)
    
    Call Chama_Tela("ImportarDadosArq", objTabelaPrecoItem)
    
'
'    If Rotina_Aviso(vbYesNo, "AVISO_IMPORTACAO_TABELAPRECO", Tabela.Text, DataVigencia.Text) = vbYes Then
'
'        lErro = CF("ImportTabelaPrecoItem_Le", colTabelaPrecoItem)
'        If lErro <> SUCESSO Then gError 131018
'
'        For Each objTabelaPrecoItem In colTabelaPrecoItem
'
'            objTabelaPrecoItem.dtDataVigencia = StrParaDate(DataVigencia.Text)
'
'            If objTabelaPrecoItem.iFilialEmpresa = 0 Then objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
'
'            objTabelaPrecoItem.iCodTabela = StrParaInt(Tabela.Text)
'
'            lErro = CF("TabelaPrecoItem_Grava", objTabelaPrecoItem)
'            If lErro <> SUCESSO Then gError 131019
'
'        Next
'
'        Call Limpa_Tela_TabelaPrecoItem_Grava
'
'    End If
'
'    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoImportarItens_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 131015
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)
        
        Case 131016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_VIGENCIA_NAO_PREENCHIDA", gErr)
        
        Case 131017
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_VIGENCIA_MENOR_DATA_ATUAL", gErr, Date)

        Case 131018 To 131019

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158286)

    End Select

    Exit Sub

End Sub
'###################################################################

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 17998

    'Limpa a tela
    Call Limpa_Tela_TabelaPrecoItem

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 17998

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158287)

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

    dtDataVigenciaAnt = DATA_NULA
    
    Set objEventoItemTabela = New AdmEvento
    Set objEventoProduto = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then Error 28005

    'Chama Carrega_TabelaPreco
    lErro = Carrega_TabelaPreco(colCodigo)
    If lErro <> SUCESSO Then Error 28149

    If Tabela.ListCount > 0 Then Tabela.ListIndex = 0

'''    'Inicializa a Lista de Produtos
'''    lErro = CF("Carga_Arvore_Produto_Venda",Produtos.Nodes)
'''    If lErro <> SUCESSO Then Error 28006

    Valor.Format = gobjFAT.sFormatoPrecoUnitario
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 28005, 28149

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158288)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Carrega_TabelaPreco(colCodigo As Collection) As Long
'Carrega a ComboBox Tabela

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Carrega_TabelaPreco

    'Preenche a ComboBox com  os Tipos de Documentos existentes no BD
    lErro = CF("TabelasPreco_Le_Codigos", colCodigo)
    If lErro <> SUCESSO Then Error 17993

    For iIndice = 1 To colCodigo.Count

        'Preenche a ComboBox Tabela com os objetos da colecao colTabelaPreco
        Tabela.AddItem colCodigo(iIndice)
        Tabela.ItemData(Tabela.NewIndex) = colCodigo(iIndice)
    Next

    Carrega_TabelaPreco = SUCESSO

    Exit Function

Erro_Carrega_TabelaPreco:

    Carrega_TabelaPreco = Err

    Select Case Err

        Case 17993

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158289)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoItemTabela = Nothing
    Set objEventoProduto = Nothing
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

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
            lErro = CF("Produto_Critica_Filial2", sProduto, objProduto, iProdutoPreenchido)
            If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then Error 58486
    
            If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then Error 58487
            
            'Se a Tabela de
            If Tabela.ListIndex < 0 Then Error 58488
                
            'Preenche os dados do item da Tabela na Tela com a Data de Vigência escolhida
            lErro = Carrega_Dados_Tabela_DataVigencia(objProduto)
            If lErro <> SUCESSO Then Error 58489
                
        End If
    
    End If
    
    dtDataVigenciaAnt = StrParaDate(DataVigencia.Text)

    Exit Sub
    
Erro_DataVigencia_Click:

    Select Case Err
        
        Case 58487, 58488 'Produto ou Tabela não Preenchidos
            
        Case 58486, 58489 'Tratados nas rotinas chamadas
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158290)

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
    
    'Lê o Valor no Bd
    lErro = CF("TabelaPrecoItem_Le_DataVigente", objTabelaPrecoItem)
    If lErro <> SUCESSO And lErro <> 58503 Then Error 58490
    
    If lErro = 58503 Then Error 58491
    
    If objTabelaPrecoItem.dPreco <> 0 Then
        Valor.Text = Format(objTabelaPrecoItem.dPreco, Valor.Format)
    Else
        Valor.Text = ""
    End If
    
    PercDesconto.Text = Format(objTabelaPrecoItem.dPercDesconto * 100, "FIXED")
    PrecoComDesconto.Caption = Format(objTabelaPrecoItem.dPrecoComDesconto, gobjFAT.sFormatoPrecoUnitario)
    
    Observacao.Text = objTabelaPrecoItem.sObservacao
    
    Call Trata_Desconto
    
    iAlterado = 0
   
    Carrega_Dados_Tabela_DataVigencia = SUCESSO
    
    Exit Function
    
Erro_Carrega_Dados_Tabela_DataVigencia:

    Carrega_Dados_Tabela_DataVigencia = Err
    
    Select Case Err
    
        Case 58490
        
        Case 58491
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECOITEM_INEXISTENTE2", Err, objTabelaPrecoItem.iCodTabela, objTabelaPrecoItem.sCodProduto, objTabelaPrecoItem.dtDataVigencia)
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158291)

    End Select

    Exit Function
    
End Function

Private Sub objEventoItemTabela_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTabelaPrecoItem As ClassTabelaPrecoItem

On Error GoTo Erro_objEventoItemTabela_evSelecao

    Set objTabelaPrecoItem = obj1

    'Traz item de Tabela de Preço para a tela
    lErro = Traz_TabelaPrecoItem_Tela(objTabelaPrecoItem)
    If lErro <> SUCESSO Then Error 49865
        
    Me.Show

    Exit Sub

Erro_objEventoItemTabela_evSelecao:

    Select Case Err

        Case 49865
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158292)

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
Dim objTabelaPrecoItem1 As New ClassTabelaPrecoItem
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String
Dim iProdutoPreenchido As Integer, objTabelaPreco As New ClassTabelaPreco

On Error GoTo Erro_Produto_Validate

    If iProdutoAlterado = 1 Then
    
        If Len(Trim(Produto.ClipText)) > 0 Then
        
            sProduto = Produto.Text
    
            'Critica o formato do Produto e se existe no BD
            lErro = CF("Produto_Critica_Filial2", sProduto, objProduto, iProdutoPreenchido)
            If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then Error 28111
    
            If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then Error 27204
            
            If lErro = 51381 Then Error 28113
            
            'Preenche ProdutoDescricao com Descrição do Produto
            DescricaoProduto.Caption = objProduto.sDescricao
            Call CF2(Me, "TabPreco_Produto_ExibeInfo", objProduto.sCodigo, objProduto.sCor)
    
            Valor.Text = ""
            ValorEmpresa.Caption = ""
            
            'Limpa a Compo de Data de Vigencia
            DataVigencia.Clear
            
            UnidadeMedida.Caption = ""
            
            'Verifica se Tabela está preenchida
            If Len(Trim(Tabela.Text)) > 0 Then
                            
                'Carrega a Combo de Data de Vigência
                lErro = Carrega_Combo_DataVigencia(objProduto)
                If lErro <> SUCESSO And lErro <> 58497 Then Error 58495
                            
                'Carrega os dados na Tela
                lErro = Carrega_Dados_TabelaPreco(objProduto)
                If lErro <> SUCESSO Then Error 28151
                
                objTabelaPreco.iCodigo = CInt(Tabela.Text)
                lErro = CF("TabelaPreco_le", objTabelaPreco)
                If lErro <> SUCESSO And lErro <> 28004 Then Error 58495
                If lErro = SUCESSO Then
                    If objTabelaPreco.iTipo = TABELA_PRECO_TIPO_COMPRA Then UnidadeMedida.Caption = objProduto.sSiglaUMCompra
                End If

            End If
    
            'Mostra a unidade de medida na tela
            If UnidadeMedida.Caption = "" Then
                
                UnidadeMedida.Caption = objProduto.sSiglaUMVenda
            
                If objProduto.iFaturamento = PRODUTO_NAO_VENDAVEL Then Error 62091
            
            End If
            
        Else
            'Limpa DescricaoProduto
            DescricaoProduto.Caption = ""
            Call CF2(Me, "TabPreco_Produto_ExibeInfo", "")
            
            'Limpa unidade de medida do produto
            UnidadeMedida.Caption = ""

        End If
        
        iProdutoAlterado = 0
        
    End If
    
    Call Trata_Desconto
    
    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case Err

        Case 28111

        Case 27204, 28113 'Não encontrou Produto no BD

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

        Case 28151, 58495
        
        Case 62091
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_VENDIDO", Err, Produto.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158293)

    End Select

    Exit Sub

End Sub

Function Carrega_Combo_DataVigencia(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim colDatasVigentes As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Combo_DataVigencia
    
    objTabelaPrecoItem.sCodProduto = objProduto.sCodigo
    objTabelaPrecoItem.iCodTabela = CInt(Tabela.List(Tabela.ListIndex))
    objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
    
    'Lê as Data de Vigência para o par (Produto e Tabela)
    lErro = CF("TabelaDePrecoItens_Le_DatasVigentes", objTabelaPrecoItem, colDatasVigentes)
    If lErro <> SUCESSO And lErro <> 58483 Then Error 58496
    
    'Se nao encontrou ---> sai
    If lErro = 58483 Then Error 58497
    
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
        
        Case 58496
        
        Case 58497
            Valor.Text = ""
            DataVigencia.Text = Format(gdtDataAtual, "dd/mm/yyyy")
            Call Trata_Desconto
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158294)

    End Select

    Exit Function
    
End Function

Public Sub ProdutoLabel_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto, objTabelaPreco As New ClassTabelaPreco
Dim colSelecao As New Collection, bCompra As Boolean

On Error GoTo Erro_ProdutoLabel_Click

    'Verifica se o produto foi preenchido
    If Len(Produto.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82673

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
    
        Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProduto)

    Else
    
        Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)
    
    End If

    Exit Sub

Erro_ProdutoLabel_Click:

    Select Case gErr

        Case 82673

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158295)

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
Dim objTabelaPrecoItem1 As New ClassTabelaPrecoItem
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Error_Tabela_Click

    'Verifica se foi preenchida a ComboBox Tabela
    If Tabela.ListIndex <> -1 Then

        objTabelaPreco.iCodigo = CInt(Tabela.Text)

        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then Error 28025

        If lErro = 28004 Then Error 28026

        DescricaoTabela.Caption = objTabelaPreco.sDescricao

        If Len(Trim(Produto.ClipText)) > 0 Then
         
            sProduto = Produto.Text
        
            'Critica o Produto
            lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
            If lErro <> SUCESSO And lErro <> 25041 Then Error 28246
            
            If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then Error 27205
              
            'Carrega a Combo de Data de Vigência
            lErro = Carrega_Combo_DataVigencia(objProduto)
            If lErro <> SUCESSO And lErro <> 58497 Then Error 58499
      
            'Preeche os Dados da Tabela
            lErro = Carrega_Dados_TabelaPreco(objProduto)
            If lErro <> SUCESSO Then Error 28247
        
        End If
    
    End If

    iAlterado = 0

    Exit Sub

Error_Tabela_Click:

    Select Case Err

        Case 28025, 28246, 28247, 27205, 58499

        Case 28026
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_INEXISTENTE", Err, objTabelaPreco.iCodigo)
            Tabela.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158296)

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
        If lErro <> SUCESSO Then Error 49864

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
            
    'Abre a Tela que Lista as Tabelas de Preço
    Call Chama_Tela("TabPrecoItemLista", colSelecao, objTabelaPrecoItem, objEventoItemTabela)

    Exit Sub

Erro_TabelaLabel_Click:

    Select Case Err

        Case 49864

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 158297)

    End Select

    Exit Sub

End Sub

Public Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Sub Limpa_Tela_TabelaPrecoItem()
'Limpa a Tela TabelarecoItem

Dim lErro As Long

    DescricaoProduto.Caption = ""
    Call CF2(Me, "TabPreco_Produto_ExibeInfo", "")
    UnidadeMedida.Caption = ""
    ValorEmpresa.Caption = ""
    DescricaoTabela.Caption = ""
    PrecoComDesconto.Caption = ""
    
    dtDataVigenciaAnt = DATA_NULA

    TabelaDefault.Value = vbUnchecked
    
    'Funcao generica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    Tabela.ListIndex = -1
    
    DataVigencia.Clear
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

End Sub

Sub Limpa_Tela_TabelaPrecoItem_Grava()
'Limpa a Tela TabelaPrecoItem quando esta na gravação

Dim lErro As Long

    DescricaoProduto.Caption = ""
    Call CF2(Me, "TabPreco_Produto_ExibeInfo", "")
    UnidadeMedida.Caption = ""
    ValorEmpresa.Caption = ""
    PrecoComDesconto.Caption = ""

    TabelaDefault.Value = vbUnchecked
    
    dtDataVigenciaAnt = DATA_NULA
    
    DataVigencia.Clear

    'Funcao generica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

End Sub

Public Sub Valor_GotFocus()

Dim objValor As Object

    Set objValor = Valor

    Call CF("TabelaPrecoItem_Valor_GotFocus", objValor)

End Sub

Public Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(Valor.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_Positivo_Critica(Valor.Text)
    If lErro <> SUCESSO Then Error 17999

    Valor.Text = Format(Valor.Text, Valor.Format)
    
    Call Trata_Desconto

    Exit Sub

Erro_Valor_Validate:

    Cancel = True


    Select Case Err

        Case 17999

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158298)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTabelaPrecoItem As ClassTabelaPrecoItem) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um item selecionado, exibir seus dados
    If Not (objTabelaPrecoItem Is Nothing) Then

        'Chama Traz_TabelaPrecoItem_Tela
        lErro = Traz_TabelaPrecoItem_Tela(objTabelaPrecoItem)
        If lErro <> SUCESSO Then Error 28019

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 28019

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158299)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim colCodFilial As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TabelaPrecoItemProduto"

    'Lê os dados da Tela Tabela Preço Item
    lErro = Move_Tela_Memoria(objTabelaPrecoItem)
    If lErro <> SUCESSO Then Error 28021

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "FilialEmpresa", giFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "CodTabela", objTabelaPrecoItem.iCodTabela, 0, "CodTabela"
    colCampoValor.Add "CodProduto", objTabelaPrecoItem.sCodProduto, STRING_PRODUTO, "CodProduto"

    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 28021

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158300)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem

On Error GoTo Erro_Tela_Preenche

    objTabelaPrecoItem.iCodTabela = colCampoValor.Item("CodTabela").vValor
    objTabelaPrecoItem.sCodProduto = colCampoValor.Item("CodProduto").vValor

    If (objTabelaPrecoItem.iCodTabela <> 0) And (objTabelaPrecoItem.sCodProduto <> "") Then

        'Traz os dados para tela
        lErro = Traz_TabelaPrecoItem_Tela(objTabelaPrecoItem)
        If lErro <> SUCESSO Then Error 28023

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 28023

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158301)

    End Select

    Exit Sub

End Sub

Private Function Traz_TabelaPrecoItem_Tela(objTabelaPrecoItem As ClassTabelaPrecoItem) As Long
'Traz os dados da Tabela Preço Item para a Tela

Dim lErro As Long
Dim sCodigo As String
Dim sProdutoEnxuto As String
Dim objTabelaPreco As New ClassTabelaPreco
Dim objTabelaPrecoItem1 As New ClassTabelaPrecoItem

On Error GoTo Erro_Traz_TabelaPrecoItem_Tela
    
    'Seleciona a Tabela na Combo
    Call Combo_Seleciona_ItemData(Tabela, objTabelaPrecoItem.iCodTabela)
    
    If Tabela.ListIndex = -1 Then Error 28180
    
    'Preenche o Produto
    lErro = CF("Traz_Produto_MaskEd", objTabelaPrecoItem.sCodProduto, Produto, DescricaoProduto)
    If lErro <> SUCESSO Then Error 28178

    Call Produto_Validate(bSGECancelDummy)

    'Observacao.Text = objTabelaPrecoItem.sObservacao
    
    iAlterado = 0

    Traz_TabelaPrecoItem_Tela = SUCESSO

    Exit Function

Erro_Traz_TabelaPrecoItem_Tela:

    Traz_TabelaPrecoItem_Tela = Err

    Select Case Err
    
        Case 28178

        Case 28180
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECOITEM_INEXISTENTE", Err, objTabelaPrecoItem.iCodTabela)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158302)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento de tabela
    If Len(Trim(Tabela.Text)) = 0 Then gError 28064

    'Verifica preenchimento de produto
    If Len(Trim(Produto.ClipText)) = 0 Then gError 28065

    'Verifica preenchimento de valor
    If Len(Trim(Valor.Text)) = 0 Then gError 28066
    
    'Verifica o Preenchimento da Data de Vigencia
    If Len(Trim(DataVigencia.Text)) = 0 Then gError 58471
    
    'Verifica se a data de Vigencia é maior que date
    If CDate(DataVigencia) < Date Then gError 58472
        
    'Chama Move_Tela_Memoria
    lErro = Move_Tela_Memoria(objTabelaPrecoItem)
    If lErro <> SUCESSO Then gError 28067

    lErro = Trata_Alteracao(objTabelaPrecoItem, objTabelaPrecoItem.iCodTabela, objTabelaPrecoItem.iFilialEmpresa, objTabelaPrecoItem.sCodProduto, objTabelaPrecoItem.dtDataVigencia)
    If lErro <> SUCESSO Then gError 32330

    sProduto = Produto.Text

    'Critica o formato do Produto
    lErro = CF("Produto_Critica2", sProduto, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 20851

    If lErro = 25043 And objProduto.iKitVendaComp <> MARCADO And Len(objProduto.sGrade) = 0 Then
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_PRODUTO_GERENCIAL_TABPRECO", Produto.Text)

        If vbMsgRes = vbNo Then gError 20852

    End If

    lErro = CF2(Me, "TABPRE_Grava_Customiza", objTabelaPrecoItem, objProduto)
    If lErro <> SUCESSO Then gError 20851

    'Chama TabelaPrecoItem_Grava
    lErro = CF("TabelaPrecoItem_Grava", objTabelaPrecoItem)
    If lErro <> SUCESSO Then gError 28068

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 20851, 20852, 28067, 28068, 28102, 28191, 32330

        Case 28064
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)

        Case 28065
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 28066
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO1", gErr)

        Case 28103
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 58471
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_VIGENCIA_NAO_PREENCHIDA", gErr)
        
        Case 58472
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_VIGENCIA_MENOR_DATA_ATUAL", gErr, Date)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158303)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objTabelaPrecoItem As ClassTabelaPrecoItem) As Long
'Move os dados da Tela para objTabelaPrecoItem

Dim lErro As Long
Dim objTabelaPreco As New ClassTabelaPreco
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
    
    If Len(Trim(Tabela.Text)) > 0 Then objTabelaPrecoItem.iCodTabela = CInt(Tabela.Text)

    If Len(Trim(Produto.ClipText)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 28191
        objTabelaPrecoItem.sCodProduto = sProdutoFormatado
    
    End If
    
    objTabelaPrecoItem.iTabelaDefault = TabelaDefault.Value

    'Preenche objTabelaPrecoItem com preço e DataVigencia
    If Len(Trim(Valor.ClipText)) > 0 Then objTabelaPrecoItem.dPreco = CDbl(Valor.ClipText)
    
    'Move a Data de Vigência para a Memoria
    If Len(Trim(DataVigencia.Text)) > 0 Then objTabelaPrecoItem.dtDataVigencia = CDate(DataVigencia.Text)
    
    objTabelaPrecoItem.sObservacao = Observacao.Text
    
    objTabelaPrecoItem.dPercDesconto = StrParaDbl(PercDesconto.Text) / 100
    objTabelaPrecoItem.dPrecoComDesconto = StrParaDbl(PrecoComDesconto.Caption)
    
    Set objTabelaPrecoItem.objTela = Me

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 28100, 28191

        Case 28101
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_INEXISTENTE", Err, objTabelaPreco.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158304)

    End Select

    Exit Function

End Function

Private Function Carrega_Dados_TabelaPreco(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim objTabelaPrecoItem1 As New ClassTabelaPrecoItem
Dim objProdutosFilial As New ClassProdutoFilial

On Error GoTo Erro_Carrega_Dados_TabelaPreco

    'Preenche objTabelaPrecoItem
    objTabelaPrecoItem.iCodTabela = CInt(Tabela.Text)
    objTabelaPrecoItem.sCodProduto = objProduto.sCodigo
    objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
    
    'Lê TabelaPrecoItem
    lErro = CF("TabelaPrecoItem_Le1", objTabelaPrecoItem)
    If lErro <> SUCESSO And lErro <> 126784 Then Error 28244
    
    'Se encontrou o Ítem mostra o Valor na tela
    If lErro <> 126784 Then
        If objTabelaPrecoItem.dPreco <> 0 Then
            Valor.Text = Format(objTabelaPrecoItem.dPreco, Valor.Format)
        Else
            Valor.Text = ""
        End If
        
        'Preenche a Data de Vigência
        If objTabelaPrecoItem.dtDataVigencia <> DATA_NULA Then
            DataVigencia.Text = Format(objTabelaPrecoItem.dtDataVigencia, "dd/mm/yyyy")
            Call Combo_Item_Igual(DataVigencia)
        End If
        
        Observacao.Text = objTabelaPrecoItem.sObservacao
        
        PercDesconto.Text = Format(objTabelaPrecoItem.dPercDesconto * 100, "FIXED")
        PrecoComDesconto.Caption = Format(objTabelaPrecoItem.dPrecoComDesconto, gobjFAT.sFormatoPrecoUnitario)
    
    Else
        
        'Se não encontrou o Ítem limpa o campo Valor
        Valor.Text = ""
        DataVigencia.Text = Format(gdtDataAtual, "dd/mm/yyyy")
        Observacao.Text = ""
        
    End If
    
    dtDataVigenciaAnt = objTabelaPrecoItem.dtDataVigencia
    
    'Preenche objTabelaPrecoItem1
    objTabelaPrecoItem1.iCodTabela = CInt(Tabela.Text)
    objTabelaPrecoItem1.sCodProduto = objProduto.sCodigo
    objTabelaPrecoItem1.iFilialEmpresa = EMPRESA_TODA

    'Lê TabelaPrecoItem
    lErro = CF("TabelaPrecoItem_Le1", objTabelaPrecoItem1)
    If lErro <> SUCESSO And lErro <> 126784 Then Error 28245

    'Se encontrou o Ítem mostra o ValorEmpresa na tela
    If lErro <> 126784 Then
        If objTabelaPrecoItem1.dPreco <> 0 Then
            ValorEmpresa.Caption = Format(objTabelaPrecoItem1.dPreco, Valor.Format)
        Else
            ValorEmpresa.Caption = ""
        End If
    Else
        'Limpa campo ValorEmpresa
        ValorEmpresa.Caption = ""
    End If

    objProdutosFilial.iFilialEmpresa = giFilialEmpresa
    objProdutosFilial.sProduto = objProduto.sCodigo
    
    lErro = CF("ProdutoFilial_Le", objProdutosFilial)
    If lErro <> SUCESSO And lErro <> 28261 Then Error 62092
    
    If objProduto.iKitVendaComp <> MARCADO Then
        If lErro <> SUCESSO Then Error 62093
    End If
    
    If objProdutosFilial.iTabelaPreco = objTabelaPrecoItem.iCodTabela Then
        TabelaDefault.Value = vbChecked
    Else
        TabelaDefault.Value = vbUnchecked
    End If
    
    Call Trata_Desconto

    Carrega_Dados_TabelaPreco = SUCESSO
    
    Exit Function
    
Erro_Carrega_Dados_TabelaPreco:

    Carrega_Dados_TabelaPreco = Err
    
    Select Case Err
    
        Case 28244, 28245, 62092
        
        Case 62093
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FILIAL_NAO_CADASTRADO", Err, objProdutosFilial.sProduto, objProdutosFilial.iFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158305)

    End Select

    Exit Function

End Function

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
    If lErro <> SUCESSO Then Error 58475
    
    If dtDataVigenciaAnt <> StrParaDate(DataVigencia.Text) Then Call DataVigencia_Click
    
    Exit Sub

Erro_DataVigencia_Validate:

    Cancel = True


    Select Case Err

        'se houve erro de crítica, segura o foco
        Case 58475

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158306)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ITENS_TABELA_PRECO
    Set Form_Load_Ocx = Me
    Caption = "Itens da Tabela de Preço"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TabelaPrecoItem"
    
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
            Call ProdutoLabel_Click
        End If
    
    End If

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82674

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82675
    
    Produto.PromptInclude = False
    Produto.Text = objProduto.sCodigo
    Produto.PromptInclude = True
        
    Me.Show

    'Janaina
    Call Produto_Validate(bSGECancelDummy)
    'Janaina

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 82674

        Case 82675
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158307)

    End Select

    Exit Sub

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
Function TabPreco_Produto_ExibeInfo(ByVal objCT As Object, ByVal sProduto As String, Optional ByVal sCor As String = "") As Long
    TabPreco_Produto_ExibeInfo = SUCESSO
End Function

Public Sub Observacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor, Optional objContexto As Object) As Long

Dim lErro As Long
Dim objTabelaPrecoItem As ClassTabelaPrecoItem
Dim sTexto As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
        
                        
        Case ALTTABPRECO
            Set objTabelaPrecoItem = objContexto
            
            If objTabelaPrecoItem.dPreco <> objTabelaPrecoItem.dPrecoAtual Then
                objMnemonicoValor.colValor.Add 1
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
            
        Case ALTTABPRECODESC
            
            Set objTabelaPrecoItem = objContexto
            
            sTexto = "O produto " & objTabelaPrecoItem.sCodProduto & " teve seu preço alterado na tabela " & _
            CStr(objTabelaPrecoItem.iCodTabela) & " de R$ " & Format(objTabelaPrecoItem.dPrecoAtual, Valor.Format) & _
            " para R$ " & Format(objTabelaPrecoItem.dPreco, Valor.Format) & " pelo usuário " & gsUsuario
            
            objMnemonicoValor.colValor.Add sTexto
            

        Case Else
            gError 197765

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 197765
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197766)

    End Select

    Exit Function

End Function

Function TABPRE_Grava_Customiza(ByVal objCT As Object, ByVal objTabelaPrecoItem As ClassTabelaPrecoItem, ByVal objProduto As ClassProduto) As Long
    TABPRE_Grava_Customiza = SUCESSO
End Function

Public Sub PercDesconto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercDesconto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPercDesconto As Double

On Error GoTo Erro_PercDesconto_Validate

    dPercDesconto = 0

    'Verifica se o Valor está preenchido
    If Len(Trim(PercDesconto.Text)) > 0 Then
    
        'Faz a Crítica do Valor digitado
        lErro = Porcentagem_Critica(PercDesconto.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        dPercDesconto = StrParaDbl(PercDesconto.Text) / 100
        
        PercDesconto.Text = Format(dPercDesconto * 100, "FIXED")

    End If
    
    Call Trata_Desconto
           
    Exit Sub

Erro_PercDesconto_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157330)

    End Select

    Exit Sub

End Sub

Private Sub Trata_Desconto()
Dim dPercDesconto As Double, dPreco As Double, dDesconto As Double
    dPercDesconto = StrParaDbl(PercDesconto.Text) / 100
    dPreco = StrParaDbl(Valor.Text)
    dDesconto = StrParaDbl(Format(dPreco * dPercDesconto, gobjFAT.sFormatoPrecoUnitario))
    PrecoComDesconto.Caption = Format(dPreco - dDesconto, gobjFAT.sFormatoPrecoUnitario)
End Sub
