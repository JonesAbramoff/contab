VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPCPGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''*** Daqui p/ baixo rotinas referente a Previsão de Vendas x Previsão de Consumo ******
''???? cadastrar esta função no dic
Public Function PrevVenda_ReqMat_Calcula(ByVal sProd_I As String, ByVal sProd_Fim As String, ByVal iMes As Integer, ByVal iAno As Integer, ByVal sPrevisao As String, ByVal sversao As String) As Long
'Faz a filtragem dos Parâmetros que serão utilizados no calculo da Previsão de Consumo.

Dim lErro As Long
Dim alComando(0) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim objPrevVMensal As ClassPrevVendaMensal
Dim tPrevVMensal As typePrevVendaMensal
Dim colMateriais As New Collection
Dim sSQL As String
Dim iFilialEmpresa As Integer

On Error GoTo Erro_PrevVenda_ReqMat_Calcula

    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 103210

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 103187
    Next

    With tPrevVMensal

        .sCodigo = String(STRING_PREVVENDA_CODIGO, 0)
        .sProduto = String(STRING_PRODUTO, 0)

        sSQL = "SELECT Produto, SUM(Quantidade" & iMes & ") FROM PrevVendaMensal WHERE Codigo = ? AND FilialEmpresa = ? AND Ano = ? " & _
        "Group BY Produto ORDER BY Produto"

        'Seleciona os dados de PrevVendaMensal de acordo com o Código, FilialEmpresa e Ano recebidos
        lErro = Comando_Executar(alComando(0), sSQL, _
         .sProduto, .dQuantidade1, sPrevisao, giFilialEmpresa, iAno)
        If lErro <> AD_SQL_SUCESSO Then gError 103193

    End With

    'Tenta buscar o primeiro elemento.
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103189

    Do While lErro = AD_SQL_SUCESSO
        
        Set objPrevVMensal = New ClassPrevVendaMensal

        objPrevVMensal.sProduto = tPrevVMensal.sProduto
        objPrevVMensal.dQuantidade1 = tPrevVMensal.dQuantidade1

'        If objPrevVMensal.sProduto = "0001683" Then
'            MsgBox "ok"
'        End If

'        Select Case objPrevVMensal.sProduto
'            Case "0000350", "0002062", "0002121", "0002122", "0002128", "0002173"
'                MsgBox "aaa"
'        End Select
        
        'Prepara a gravação dos elementos lidos e dispara a rotina de gravação em PrevVendaPrevConsumo.
        lErro = PrevVendaPrevConsumo_Grava(objPrevVMensal, iMes, iAno, sversao, colMateriais)
        If lErro <> SUCESSO Then gError 103204
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103192
            
    Loop
    
    lErro = PrevVendaPrevConsumo_Grava_BD(colMateriais, iAno, iMes, sPrevisao)
    If lErro <> SUCESSO Then gError 103223

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 103217

    PrevVenda_ReqMat_Calcula = SUCESSO

    Exit Function

Erro_PrevVenda_ReqMat_Calcula:

    PrevVenda_ReqMat_Calcula = gErr

    Select Case gErr

        Case 103210
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 103217
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 103189, 103192, 103193
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVVENDAMENSAL", gErr, sPrevisao)

        Case 103187
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103204, 103223

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Call Transacao_Rollback

    Exit Function

End Function

'???? cadastrar esta função no dic
Public Function PrevVendaPrevConsumo_Grava(ByVal objPrevVMensal As ClassPrevVendaMensal, ByVal iMes As Integer, ByVal iAno As Integer, sversao As String, ByVal colMateriais As Collection) As Long
'Prepara a gravação dos elementos lidos e dispara a rotina de gravação em PrevVendaPrevConsumo.

Dim lErro As Long
Dim objPrevConsumo As New ClassPrevVendaPrevConsumo
Dim objProdutoKit As New ClassProdutoKit
Dim objKit As ClassKit
Dim objProduto As New ClassProduto
Dim dFatorConversao As Double

On Error GoTo Erro_PrevVendaPrevConsumo_Grava

    objProduto.sCodigo = objPrevVMensal.sProduto

    'Lê o produto para descobrir as unidades de medidas associadas
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 103206

    'Descobre o fator de conversao da UM de Venda p/UM de estoque do produto raiz
    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMVenda, objProduto.sSiglaUMEstoque, dFatorConversao)
    If lErro <> SUCESSO Then gError 103207

    'Converte a quantidade na UM de venda para a quantidade equivalente na UM de estoque.
    objPrevConsumo.dQuantPrevInsumo = objPrevVMensal.dQuantidade1 * dFatorConversao
    objPrevConsumo.dQuantPrevVenda = objPrevConsumo.dQuantPrevInsumo
        
    'Carrega o obj com a Versão do produto
    objPrevConsumo.sversao = sversao
    objPrevConsumo.sProduto = objPrevVMensal.sProduto
    objPrevConsumo.sUMEstoque = objProduto.sSiglaUMEstoque

    lErro = Calcula_Insumos_PrevConsumo(colMateriais, objPrevConsumo, 0)
    If lErro <> SUCESSO Then gError 103209

    PrevVendaPrevConsumo_Grava = SUCESSO

    Exit Function

Erro_PrevVendaPrevConsumo_Grava:

    PrevVendaPrevConsumo_Grava = gErr

    Select Case gErr

        Case 103196, 103205, 103206, 103207, 103208, 103209

        Case 103195
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Function Calcula_Insumos_PrevConsumo(ByVal colMateriais As Collection, ByVal objPrevConsumo As ClassPrevVendaPrevConsumo, ByVal iNiveis As Integer) As Long
'Calcula de forma recursiva todos os componentes necessários para a fabricação de
'um Kit até último nivel.

Dim lErro As Long
Dim objPrevConsumo1 As ClassPrevVendaPrevConsumo
Dim objKit As New ClassKit
Dim objKitAux As New ClassKit
Dim bEncontrou As Boolean
Dim objProdutoKit As New ClassProdutoKit
Dim objProdutoKitAux As New ClassProdutoKit
Dim objProduto As New ClassProduto
Dim dQuantidadeFilho As Double
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim iNiveisAux As Integer
Dim dFatorConversao As Double
Dim dQuantidade As Double
Dim dFatorConversaoProdutoRaiz As Double

On Error GoTo Erro_Calcula_Insumos_PrevConsumo

    'Se a quantidade de chegou a 1 Kg ou já desceu na recursividade 25 instâncias
    If (objPrevConsumo.dQuantPrevInsumo <= 1 And iNiveis > 0) Or iNiveis >= 20 Then
        Calcula_Insumos_PrevConsumo = SUCESSO
        Exit Function

    End If

'    Select Case objPrevConsumo.sProduto
'        Case "A000010"
'            MsgBox "bbb"
'    End Select
    
    'verifica se o produto está na coleção de produtos
    For Each objPrevConsumo1 In colMateriais

        'se o Produto existe
        If objPrevConsumo1.sProduto = objPrevConsumo.sProduto Then

            objPrevConsumo1.dQuantPrevInsumo = objPrevConsumo1.dQuantPrevInsumo + objPrevConsumo.dQuantPrevInsumo
            objPrevConsumo1.dQuantPrevVenda = objPrevConsumo1.dQuantPrevVenda + objPrevConsumo.dQuantPrevVenda
            bEncontrou = True
            Exit For

        End If

    Next
    
    'Se ainda não existe na coleção de produtos
    If Not bEncontrou Then

        Set objPrevConsumo1 = objPrevConsumo
        
        objEstoqueProduto.sProduto = objPrevConsumo.sProduto
        
        'Le a qtd do produto em estoque
        lErro = CF("EstoqueProduto_ObterQTD_Filial", objEstoqueProduto, giFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 103277 Then gError 103220

        'se o produto realmente for da filial em questao
        If lErro = SUCESSO Then
        
            objPrevConsumo1.dQuantEstoque = objEstoqueProduto.dQuantDispNossa
            
            'Adiociona este produto em uma coleção global
            colMateriais.Add objPrevConsumo1

        End If

    End If

    

    objKit.sversao = objPrevConsumo.sversao
    objKit.sProdutoRaiz = objPrevConsumo.sProduto

    'tem que ler o proprio cara e os seus filhos
    lErro = CF("ProdutoKit_Le_PrimeiroNivel_Versao", objKit)
    If lErro <> SUCESSO And lErro <> 103203 Then gError 103221
    
    'Se o produto for um Kit, então leva a recursividade em diante
    If lErro = SUCESSO Then
    
        'verifica se o produto é produzido na filial que está sendo analisada
        lErro = CF("Produto_ProducaoLocal", objPrevConsumo.sProduto, giFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 103277 Then gError 106974
        
        If lErro = SUCESSO Then
    
            'percorre os elementos componentes do Kit Lido
            For Each objProdutoKit In objKit.colComponentes
            
                objProduto.sCodigo = objProdutoKit.sProduto
        
                'Lê o produto para descobrir as unidades de medidas associadas
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 103218
        
                'Descobre o fator de conversao da UM recebida p/UM de estoque do produto
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProdutoKit.sUnidadeMed, objProduto.sSiglaUMEstoque, dFatorConversao)
                If lErro <> SUCESSO Then gError 103219
            
                If objProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
                    
                    dQuantidade = (objPrevConsumo.dQuantPrevInsumo / objProdutoKit.dQuantidade * dFatorConversao) / (1 - objProdutoKit.dPercentualPerda)
                    
                Else
                    
                    Set objPrevConsumo = New ClassPrevVendaPrevConsumo
            
                    objPrevConsumo.sProduto = objProdutoKit.sProdutoRaiz
                    objPrevConsumo.sversao = objProdutoKit.sversao
            
                    'Converte a quantidade na UM de venda para a quantidade equivalente na UM de estoque.
                    objPrevConsumo.dQuantPrevInsumo = dQuantidade * (objProdutoKit.dQuantidade / (1 - objProdutoKit.dPercentualPerda) * dFatorConversao)
                    objPrevConsumo.sProduto = objProdutoKit.sProduto
                    objPrevConsumo.sUMEstoque = objProduto.sSiglaUMEstoque
                    objPrevConsumo.sversao = ""
                      
                    'Calcula de maneira recursiva a quantidade de componentes necessários.
                    lErro = Calcula_Insumos_PrevConsumo(colMateriais, objPrevConsumo, iNiveis + 1)
                    If lErro <> SUCESSO Then gError 103222
            
                End If
            
            Next
    
        End If
        
    End If

    Calcula_Insumos_PrevConsumo = SUCESSO

    Exit Function

Erro_Calcula_Insumos_PrevConsumo:

    Calcula_Insumos_PrevConsumo = gErr

    Select Case gErr

        Case 103218 To 103222, 103224, 103225, 106974

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select


    Exit Function

End Function

Public Function PrevVendaPrevConsumo_Grava_BD(ByVal colMateriais As Collection, ByVal iAno As Integer, ByVal iMes As Integer, sCodigo As String) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim objPrevVendaPrevConsumo As New ClassPrevVendaPrevConsumo
Dim dQuantidade As Double
Dim dQuantPrevVenda As Double

On Error GoTo Erro_PrevVendaPrevConsumo_Grava_BD

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 103211
    Next

    'exclui os registros que tenham sido gerados numa execucao anterior do relatorio com o mesmo input
    lErro = Comando_Executar(alComando(2), "DELETE FROM PrevVendaPrevConsumo WHERE FilialEmpresa = ? AND CodigoPrevVenda = ? AND Ano = ? AND Mes = ?", giFilialEmpresa, sCodigo, iAno, iMes)
    If lErro <> AD_SQL_SUCESSO Then gError 106759

    For Each objPrevVendaPrevConsumo In colMateriais
            
        'Verifica se o produto já existe
        lErro = Comando_ExecutarPos(alComando(0), "SELECT QuantPrevVenda FROM PrevVendaPrevConsumo WHERE Produto = ? AND FilialEmpresa = ? AND CodigoPrevVenda = ? AND Ano = ? AND Mes = ?", _
        0, dQuantPrevVenda, objPrevVendaPrevConsumo.sProduto, giFilialEmpresa, sCodigo, iAno, iMes)
        If lErro <> AD_SQL_SUCESSO Then gError 103212

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103213

        If lErro = AD_SQL_SUCESSO Then

            lErro = Comando_ExecutarPos(alComando(1), "UPDATE PrevVendaPrevConsumo SET UMEstoque = ?, QuantPrevVenda = ?, QuantPrevInsumo = ?, QuantEstoque = ?, Data = ?", _
            alComando(0), objPrevVendaPrevConsumo.sUMEstoque, objPrevVendaPrevConsumo.dQuantPrevVenda, objPrevVendaPrevConsumo.dQuantPrevInsumo, objPrevVendaPrevConsumo.dQuantEstoque, gdtDataHoje)
            If lErro <> AD_SQL_SUCESSO Then gError 103214

        Else

          With objPrevVendaPrevConsumo

                lErro = Comando_Executar(alComando(0), "INSERT INTO PrevVendaPrevConsumo(FilialEmpresa, CodigoPrevVenda, Produto, Ano, Mes, UMEstoque, QuantPrevInsumo, QuantEstoque, QuantPrevVenda, Data) VALUES (?,?,?,?,?,?,?,?,?,?)", _
                giFilialEmpresa, sCodigo, .sProduto, iAno, iMes, .sUMEstoque, .dQuantPrevInsumo, .dQuantEstoque, .dQuantPrevVenda, gdtDataHoje)
                If lErro <> AD_SQL_SUCESSO Then gError 103216

            End With

        End If

    Next

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    PrevVendaPrevConsumo_Grava_BD = SUCESSO

    Exit Function

Erro_PrevVendaPrevConsumo_Grava_BD:

    PrevVendaPrevConsumo_Grava_BD = gErr

    Select Case gErr

        Case 103211
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103212, 103213
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVENDAPREVCONSUMO", gErr)

        Case 103214
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_PREVENDAPREVCONSUMO", gErr)

        Case 103216
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PREVENDAPREVCONSUMO", gErr)

        Case 106759
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PREVENDAPREVCONSUMO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Call Transacao_Rollback

    Exit Function

End Function

Function FPreco_ConvQtdeOrcCusto(ByVal sProduto As String, dQtde As Double) As Long
'converte qtde orcada p/venda da unidade de medida de venda para a de calculo de custo p/formacao de precos
'??? nao faz nada pq as previsoes estao todas em kg que é a unidade de estoque

    FPreco_ConvQtdeOrcCusto = SUCESSO

End Function

Function CustoProd_AjustarPerda(ByVal sProduto As String, ByVal sversao As String, ByVal iFilialEmpresa As Integer, ByVal iRotinaOrigem As Integer, dPercentualPerda As Double) As Long

    If iFilialEmpresa = 3 Then
        
        Select Case iRotinaOrigem
        
            Case FORMACAO_PRECO_REL_COMP_CONSUMO
                dPercentualPerda = 0
                
            Case FORMACAO_PRECO_ROTINA_CUSTOFIXO, FORMACAO_PRECO_ROTINA_CALCPRECO, FORMACAO_PRECO_ANALISE_MARGCONTR, FORMACAO_PRECO_ROTINA_CUSTOSDIRETOS, FORMACAO_PRECO_ANALISE_MARGCONTR_REL
        
                Select Case sProduto
                    
                    Case "0000137", "0002234", "D000120", "Z000412", "Z000426", "Z000427", "Z000440", "Z000445", "Z001875", "Z001940", "Z002041", "Z002207", "Z002260", "Z003910", "Z000433", "Z000159", "Z000364", "Z000345", "Z000365", "Z000428", "Z002270", "Z000435"
                        dPercentualPerda = 0
                        
                    Case "0002287"
                        dPercentualPerda = 0 ' era -0.16279 até 21/03/05
                        
                End Select
                
        End Select
        
    End If
    
    CustoProd_AjustarPerda = SUCESSO
    
End Function

