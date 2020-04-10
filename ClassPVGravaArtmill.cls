VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPVGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function BloqueioPagtoAnt_Adiciona(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Inclui no BD um bloqueio de pagtos antecipados quando a condicao de pagamento for pagto antecipado

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_BloqueioPagtoAnt_Adiciona

    objCondicaoPagto.iCodigo = objPedidoVenda.iCondicaoPagto
    
    'Lê a Cond Pagto
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then gError 126283
    
    If objCondicaoPagto.sDescReduzida = PAGAMENTO_ANTECIPADO Then
    
        'Adiciona o bloqueio ao Pedido
        objPedidoVenda.colBloqueiosPV.Add objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, 0, BLOQUEIO_PAGTO_ANTECIPADO, gsUsuario, BLOQUEIO_AUTO_RESP, gdtDataHoje, "", "", DATA_NULA, ""
    
    End If
    
    BloqueioPagtoAnt_Adiciona = SUCESSO
     
    Exit Function
    
Erro_BloqueioPagtoAnt_Adiciona:

    BloqueioPagtoAnt_Adiciona = gErr
     
    Select Case gErr
          
        Case 126283
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175942)
     
    End Select
     
    Exit Function
    
End Function

Public Function BloqueioPagtoAnt_Altera(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objPedidoVenda_BD As ClassPedidoDeVenda) As Long
'Altera bloqueio de pagto antecipado

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim objBloqueioPV_BD As ClassBloqueioPV
Dim bPossuiBloqueioPagAnt As Boolean

On Error GoTo Erro_BloqueioPagtoAnt_Altera

    objCondicaoPagto.iCodigo = objPedidoVenda.iCondicaoPagto
    
    'Lê a Cond Pagto
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then gError 126741
    
    If objCondicaoPagto.sDescReduzida <> PAGAMENTO_ANTECIPADO Then
    
        For Each objBloqueioPV_BD In objPedidoVenda_BD.colBloqueiosPV
            If objBloqueioPV_BD.iTipoDeBloqueio = BLOQUEIO_PAGTO_ANTECIPADO Then
                lErro = CF("BloqueiosPV_Exclui_Tipo", objPedidoVenda, BLOQUEIO_PAGTO_ANTECIPADO)
                If lErro <> SUCESSO Then gError 126742
            
            End If
        Next
    
    Else
    
        bPossuiBloqueioPagAnt = False
    
        For Each objBloqueioPV_BD In objPedidoVenda_BD.colBloqueiosPV
            If objBloqueioPV_BD.iTipoDeBloqueio = BLOQUEIO_PAGTO_ANTECIPADO Then
                bPossuiBloqueioPagAnt = True
                Exit For
            End If
        Next
    
        If bPossuiBloqueioPagAnt = False Then
        
            'Adiciona o bloqueio ao Pedido
            objPedidoVenda.colBloqueiosPV.Add objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, 0, BLOQUEIO_PAGTO_ANTECIPADO, gsUsuario, BLOQUEIO_AUTO_RESP, gdtDataHoje, "", "", DATA_NULA, ""

        End If

    End If

    BloqueioPagtoAnt_Altera = SUCESSO
     
    Exit Function
    
Erro_BloqueioPagtoAnt_Altera:

    BloqueioPagtoAnt_Altera = gErr
     
    Select Case gErr
          
        Case 126741, 126742
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175943)
     
    End Select
     
    Exit Function
    
End Function

Public Function Inicializa_Campos_Obrigatorios(objUserControl As Object) As Long

    'Comentado por Wagner
    'objUserControl.Controls("Label34").ForeColor = &H80&
    objUserControl.Controls("Label35").ForeColor = &H80&
    objUserControl.Controls("TransportadoraLabel").ForeColor = &H80&
    
    Inicializa_Campos_Obrigatorios = SUCESSO
    
End Function

Public Function Artmill_PreencheDescontoICMS(ByVal sNomeReduzidoCli As String, ByVal iFilialCli As Integer, dAliquota As Double, ByVal dDescontoCli As Double) As Long

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim objEndereco As New ClassEndereco
Dim colCateg As New Collection
Dim objFilialCliCategoria As ClassFilialCliCategoria
Dim objICMSAliqExterna As New ClassICMSAliqExterna

On Error GoTo Erro_Artmill_PreencheDescontoICMS
    
    If Len(Trim(sNomeReduzidoCli)) > 0 Then

        objFilialCliente.iCodFilial = iFilialCli

        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sNomeReduzidoCli, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 126421

        objEndereco.lCodigo = objFilialCliente.lEndereco
        
        lErro = CF("Endereco_Le", objEndereco)
        If lErro <> SUCESSO And lErro <> 12309 Then gError 126422
        
        If UCase(objEndereco.sSiglaEstado) <> "SP" Then
        
            'preenche uma colecao com objClienteCategoria(s) associados a objFilialCliente
            lErro = CF("ClienteCategorias_Le_FilialCliente", objFilialCliente, colCateg)
            If lErro <> SUCESSO Then gError 126423
            
            For Each objFilialCliCategoria In colCateg
            
                If UCase(objFilialCliCategoria.sCategoria) = "DESCONTO ICMS" And UCase(objFilialCliCategoria.sItem) = "SIM" Then
                
                    objICMSAliqExterna.sSiglaOrig = "SP"
                    objICMSAliqExterna.sSiglaDest = objEndereco.sSiglaEstado
                
                    'Lê a alíquota externa de ICMS para o par Origem, Destino passados como parâmetros
                    lErro = CF("ICMSAliqExternas_Le_Origem_Destino", objICMSAliqExterna)
                    If lErro <> SUCESSO And lErro <> 94994 Then gError 126424
            
                    If lErro = 94994 Then gError 126425
            
                    dAliquota = 1 - 0.82 / (1 - objICMSAliqExterna.dAliquota)
            
                    Exit For
            
                End If
                
            Next

        End If

        dAliquota = dAliquota + dDescontoCli

    End If
    
    Artmill_PreencheDescontoICMS = SUCESSO
    
    Exit Function
     
Erro_Artmill_PreencheDescontoICMS:

    Artmill_PreencheDescontoICMS = gErr

    Select Case gErr
          
        Case 126421 To 126425
        
        Case 126425
            Call Rotina_Erro(vbOKOnly, "ERRO_ICMS_ALIQ_INTERSTADUAL_INEXISTENTE", gErr, Error)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175944)
     
    End Select
     
    Exit Function

End Function

'precodesc
Public Function Orcamento_Form_Load(ByVal objUserControl As Object, ByVal objContainer As Object) As Long

On Error GoTo Erro_Orcamento_Form_Load

    objUserControl.Controls.Add "MsMask.MaskEdBox", "PrecoDesc", objContainer
    objUserControl.Controls("PrecoDesc").Enabled = False
    objUserControl.Controls("PrecoDesc").Width = objUserControl.Controls("PrecoUnitario").Width
    objUserControl.Controls("PrecoDesc").Height = objUserControl.Controls("PrecoUnitario").Height
    objUserControl.Controls("PrecoDesc").Format = objUserControl.Controls("PrecoUnitario").Format

    objUserControl.Controls.Add "VB.TextBox", "DescricaoProduto2", objContainer
    objUserControl.Controls("DescricaoProduto2").Enabled = False
    objUserControl.Controls("DescricaoProduto2").Width = 1400 'objUserControl.Controls("DescricaoProduto").Width
    objUserControl.Controls("DescricaoProduto2").Height = objUserControl.Controls("DescricaoProduto").Height

    objUserControl.Controls.Add "VB.TextBox", "DescricaoProduto3", objContainer
    objUserControl.Controls("DescricaoProduto3").Enabled = False
    objUserControl.Controls("DescricaoProduto3").Width = 1400 'objUserControl.Controls("DescricaoProduto").Width
    objUserControl.Controls("DescricaoProduto3").Height = objUserControl.Controls("DescricaoProduto").Height

    'objUserControl.Controls("DescricaoProduto").Width = objUserControl.Controls("DescricaoProduto").Width * 2

    Orcamento_Form_Load = SUCESSO
    
    Exit Function

Erro_Orcamento_Form_Load:

    Orcamento_Form_Load = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175945)

    End Select
    
    Exit Function

End Function

'precodesc
Public Sub Orcamento_Inicializa_Grid_Itens1(ByVal objGridInt As AdmGrid)

    objGridInt.colColuna.Add ("Preço Desc.")

End Sub

'precodesc
Public Sub Orcamento_Inicializa_Grid_Itens2(ByVal objGridInt As AdmGrid, ByVal objUserControl As Object)

    objGridInt.colCampo.Add objUserControl.Controls("PrecoDesc").Name

End Sub

'precodesc
Public Sub Orcamento_Inicializa_Grid_Itens3(iIncremento As Integer)

    iIncremento = iIncremento + 1

End Sub

Public Sub Orcamento_Inicializa_Grid_Itens4(ByVal objGridInt As AdmGrid)

    objGridInt.colColuna.Add ("Cor")
    objGridInt.colColuna.Add ("Detalhe")

End Sub

Public Sub Orcamento_Inicializa_Grid_Itens5(ByVal objGridInt As AdmGrid, ByVal objUserControl As Object)

    objGridInt.colCampo.Add objUserControl.Controls("DescricaoProduto2").Name
    objGridInt.colCampo.Add objUserControl.Controls("DescricaoProduto3").Name

End Sub

Public Sub Orcamento_Inicializa_Grid_Itens6(iIncremento As Integer)

    iIncremento = iIncremento + 2

End Sub

'precodesc
Public Function Carrega_Grid_Itens_PrecoDesc(ByVal objGridItens As Object, ByVal iIndice As Integer, ByVal iGrid_PrecoDesc_Col As Integer, ByVal sPrecoDesc As String, ByVal sProduto As String, ByVal iGrid_DescProduto_Col As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Carrega_Grid_Itens_PrecoDesc

    objGridItens.TextMatrix(iIndice, iGrid_PrecoDesc_Col) = sPrecoDesc

    lErro = Carrega_Grid_Itens_CorDetalhe(objGridItens, sProduto, iIndice, iGrid_DescProduto_Col)
    If lErro <> SUCESSO Then gError 126721

    Carrega_Grid_Itens_PrecoDesc = SUCESSO

    Exit Function

Erro_Carrega_Grid_Itens_PrecoDesc:

    Carrega_Grid_Itens_PrecoDesc = gErr
    
    Select Case gErr
          
        Case 126721
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175946)
     
    End Select
     
    Exit Function

End Function

Private Function Carrega_Grid_Itens_CorDetalhe(ByVal objGridItens As Object, ByVal sProduto As String, ByVal iIndice As Integer, ByVal iGrid_DescProduto_Col As Integer) As Long

Dim lErro As Long, objProduto As New ClassProduto

On Error GoTo Erro_Carrega_Grid_Itens_CorDetalhe

    objProduto.sCodigo = sProduto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 126717

    'Se o produto não está cadastrado, erro
    If lErro <> SUCESSO Then gError 126718
    
    objGridItens.TextMatrix(iIndice, iGrid_DescProduto_Col + 1) = objProduto.sCor
    
    lErro = CF("Produto_Le_InfoUsu", objProduto)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 126719
    
    'Se o produto não está cadastrado, erro
    If lErro <> SUCESSO Then gError 126720
    
    objGridItens.TextMatrix(iIndice, iGrid_DescProduto_Col + 2) = objProduto.objInfoUsu.sDetalheCor

    Carrega_Grid_Itens_CorDetalhe = SUCESSO

    Exit Function

Erro_Carrega_Grid_Itens_CorDetalhe:

    Carrega_Grid_Itens_CorDetalhe = gErr
    
    Select Case gErr
          
        Case 126717 To 126720
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175947)
     
    End Select
     
    Exit Function

End Function

'precodesc
Public Sub PrecoTotal_Calcula_PrecoDesc(ByVal objGridItens As Object, ByVal iIndice As Integer, ByVal iGrid_PrecoDesc_Col As Integer, ByVal sPrecoDesc As String)

    objGridItens.TextMatrix(iIndice, iGrid_PrecoDesc_Col) = sPrecoDesc

End Sub

'precodesc
Public Sub Produto_Saida_Celula_PrecoDesc(ByVal objGridItens As Object, ByVal iIndice As Integer, ByVal iGrid_PrecoDesc_Col As Integer, ByVal dPrecoUnitario As Double, ByVal sProduto As String, ByVal iGrid_DescProduto_Col As Integer, ByVal sCliente As String, ByVal iFilialCliente As Integer, ByVal dDescontoCli As Double, ByVal iGrid_PercDesc_Col As Integer)

Dim dPercentDesc As Double
Dim lErro As Long

On Error GoTo Erro_Produto_Saida_Celula_PrecoDesc

    If Len(Trim(sCliente)) > 0 Then

        lErro = CF("Artmill_PreencheDescontoICMS", sCliente, iFilialCliente, dPercentDesc, dDescontoCli)
        If lErro <> SUCESSO Then gError 126430

        If dPercentDesc <> 0 Then objGridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(dPercentDesc, "Percent")

    End If

    objGridItens.TextMatrix(iIndice, iGrid_PrecoDesc_Col) = Format(dPrecoUnitario * (1 - dPercentDesc), "standard")

    lErro = Carrega_Grid_Itens_CorDetalhe(objGridItens, sProduto, iIndice, iGrid_DescProduto_Col)
    If lErro <> SUCESSO Then gError 126722

    Exit Sub
    
Erro_Produto_Saida_Celula_PrecoDesc:

    Select Case gErr

    Case 126722
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175948)
     
    End Select
    
    Exit Sub

End Sub

Public Function PedidoVenda_Altera_Customizacoes(objPedidoVenda As ClassPedidoDeVenda, objPedidoVenda_BD As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim colCategoriaItem As New Collection
Dim objFilialCliCategoria As ClassFilialCliCategoria
Dim objLiberacaoCredito As New ClassLiberacaoCredito
Dim objItemPV_BD As ClassItemPedido
Dim objItemPV As ClassItemPedido

On Error GoTo Erro_PedidoVenda_Altera_Customizacoes

    If objPedidoVenda.iFilialEmpresaFaturamento <> objPedidoVenda_BD.iFilialEmpresaFaturamento Then

        objFilialCliente.lCodCliente = objPedidoVenda_BD.lCliente
        objFilialCliente.iCodFilial = objPedidoVenda_BD.iFilial
    
        'Lê na tabela de FilialClienteCategorias, todas as categorias da Filial de Cliente (com o respectivo valor em cada categoria), guardando na coleção colCategoriaItem
        lErro = CF("FilialCliente_Le_Categorias", objFilialCliente, colCategoriaItem)
        If lErro <> SUCESSO Then gError 126514
        
        For Each objFilialCliCategoria In colCategoriaItem
            If UCase(objFilialCliCategoria.sCategoria) = ARTMILL_CATCLIENTE_TROCA_FILIAL_FAT And (UCase(objFilialCliCategoria.sItem) = ARTMILL_CATCLIENTE_ITEM_NAO Or UCase(objFilialCliCategoria.sItem) = ARTMILL_CATCLIENTE_ITEM_NÃO) Then gError 126506
        Next

    End If
    
    objLiberacaoCredito.sCodUsuario = gsUsuario
    
    lErro = CF("LiberacaoCredito_Le", objLiberacaoCredito)
    If lErro <> SUCESSO And lErro <> 36968 Then gError 126507
    
    'se nao encontrou o usuario como autorizado ==> ele nao pode alterar o valor total do pedido bem como os valores unitarios e de desconto
    If lErro <> SUCESSO Then
    
'        If objPedidoVenda.dValorTotal <> objPedidoVenda_BD.dValorTotal Then gError 126508
    
        For Each objItemPV_BD In objPedidoVenda_BD.colItensPedido
    
            For Each objItemPV In objPedidoVenda.colItensPedido
            
                If objItemPV_BD.sProduto = objItemPV.sProduto Then
                
                    If objItemPV_BD.dPrecoUnitario <> objItemPV.dPrecoUnitario Then gError 126509
                    
                    If objItemPV_BD.dValorDesconto <> objItemPV.dValorDesconto Then gError 126510
                    
                    Exit For
                    
                End If
            
            Next
    
        Next
    End If
    
    PedidoVenda_Altera_Customizacoes = SUCESSO
    
    Exit Function
    
Erro_PedidoVenda_Altera_Customizacoes:

    PedidoVenda_Altera_Customizacoes = gErr
    
    Select Case gErr

        Case 126506
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PODE_ALTERAR_FILIALFAT", gErr, objPedidoVenda_BD.lCliente, objPedidoVenda_BD.iFilial)
        
        Case 126507, 126514
        
        Case 126508
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_PRECO_TOTAL_ALTERADO", gErr, objPedidoVenda_BD.dValorTotal, objPedidoVenda.dValorTotal)
        
        Case 126509
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPV_PRECO_UNITARIO_ALTERADO", gErr, objItemPV.iItem, objItemPV_BD.dPrecoUnitario, objItemPV.dPrecoUnitario)
        
        Case 126510
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPV_VALOR_DESCONTO_ALTERADO", gErr, objItemPV.iItem, objItemPV_BD.dValorDesconto, objItemPV.dValorDesconto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175949)

    End Select
    
    Exit Function


End Function

Public Function PedidoVenda_Grava_Customizacoes(objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim alComando(1 To 6) As Long
Dim iIndice As Integer
Dim iAtivo As Integer

On Error GoTo Erro_PedidoVenda_Grava_Customizacoes

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 126636
    Next

    lErro = Comando_ExecutarPos(alComando(1), "SELECT Ativo FROM Clientes WHERE Codigo = ?", 0, _
        iAtivo, objPedidoVenda.lCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 126634
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126635

    If lErro = AD_SQL_SEM_DADOS Then gError 126636
    
    If iAtivo <> CLIENTE_ATIVO Then
    
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE Clientes SET Ativo = ?", alComando(1), CLIENTE_ATIVO)
        If lErro <> AD_SQL_SUCESSO Then gError 126638
    
    End If

    lErro = Comando_ExecutarPos(alComando(3), "SELECT Ativo FROM FiliaisClientes WHERE CodCliente = ? AND CodFilial = ?", 0, _
        iAtivo, objPedidoVenda.lCliente, objPedidoVenda.iFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 126639
        
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126640

    If lErro = AD_SQL_SEM_DADOS Then gError 126641
    
    If iAtivo <> CLIENTE_ATIVO Then
    
        lErro = Comando_ExecutarPos(alComando(4), "UPDATE FiliaisClientes SET Ativo = ?", alComando(3), CLIENTE_ATIVO)
        If lErro <> AD_SQL_SUCESSO Then gError 126642
    
    End If

    If objPedidoVenda.iFilialEntrega <> 0 Then

        lErro = Comando_ExecutarPos(alComando(5), "SELECT Ativo FROM FiliaisClientes WHERE CodCliente = ? AND CodFilial = ?", 0, _
            iAtivo, objPedidoVenda.lCliente, objPedidoVenda.iFilialEntrega)
        If lErro <> AD_SQL_SUCESSO Then gError 126643
            
        lErro = Comando_BuscarPrimeiro(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126644
    
        If lErro = AD_SQL_SEM_DADOS Then gError 126645
        
        If iAtivo <> CLIENTE_ATIVO Then
        
            lErro = Comando_ExecutarPos(alComando(6), "UPDATE FiliaisClientes SET Ativo = ?", alComando(5), CLIENTE_ATIVO)
            If lErro <> AD_SQL_SUCESSO Then gError 126646
        
        End If

    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    PedidoVenda_Grava_Customizacoes = SUCESSO
    
    Exit Function
    
Erro_PedidoVenda_Grava_Customizacoes:

    PedidoVenda_Grava_Customizacoes = gErr
    
    Select Case gErr

        Case 126634, 126635
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES1", gErr)

        Case 126636
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 126637
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objPedidoVenda.lCliente)

        Case 126638
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_CLIENTES", gErr, objPedidoVenda.lCliente)
        
        Case 126639, 126640, 126643, 126644
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", gErr, objPedidoVenda.lCliente, objPedidoVenda.iFilial)

        Case 126641
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA2", gErr, objPedidoVenda.iFilial, objPedidoVenda.lCliente)

        Case 126642
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_FILIALCLIENTE1", gErr, objPedidoVenda.lCliente, objPedidoVenda.iFilial)
        
        Case 126645
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA2", gErr, objPedidoVenda.iFilialEntrega, objPedidoVenda.lCliente)
        
        Case 126646
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_FILIALCLIENTE1", gErr, objPedidoVenda.lCliente, objPedidoVenda.iFilialEntrega)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 126647)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function


