VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim lComando As Long
Dim sProdutoAnt As String
Dim iFilialEmpresaAnt As Integer
Dim dValorAnt As Double

'para calculo de estoque numa data
Private sProdutoAntInv As String
Private iInvAlmoxarifadoAntInv As Integer
Private dtDataAntInv As Date
Private dtDataFimAnt As Date


Private dQuantidadeNossaInv As Double
Private dQuantidadeEmTerc As Double
Private dQuantidadeDeTerc As Double
Private dQuantMinimaAnt As Double

Function ObterValorProduto(dValor As Double, sProduto As String, iFilialEmpresa As Integer, iAlmoxarifado As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_ObterValorProduto

    If sProduto <> sProdutoAnt Or iFilialEmpresaAnt <> iFilialEmpresa Then

        'Le o Valor em Estoque para o Produto
        lErro = Retorna_ValorEstoque_Atual(dValor, sProduto, iFilialEmpresa, iAlmoxarifado)
        If lErro <> SUCESSO Then gError 78107

        sProdutoAnt = sProduto
        iFilialEmpresaAnt = iFilialEmpresa
        dValorAnt = dValor

    End If

    dValor = dValorAnt

    Exit Function

    ObterValorProduto = SUCESSO

Erro_ObterValorProduto:

    ObterValorProduto = gErr

    Select Case gErr

        Case 78107

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153163)

    End Select

    Exit Function

End Function

Function Calcula_Valor_Custo(sProduto As String, objEstoqueProdutoSUM As ClassEstoqueProduto, dCusto As Double, dCustoConsig As Double, dCustoDemo As Double, dCustoConserto As Double, dCustoOutras As Double, dCustoBenef As Double, dCustoConsig3 As Double, dCustoDemo3 As Double, dCustoConserto3 As Double, dCustoOutras3 As Double, dCustoBenef3 As Double, iFilialEmpresa As Integer) As Long
'devolve o custo do produto e dos Escaninhos

Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Calcula_Valor_Custo
    
    objProduto.sCodigo = sProduto
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO Then gError 78108
    
    'O valor da diferenca vai depender se é um produto Produzido ou Compravel
    If objProduto.iApropriacaoCusto = APROPR_CUSTO_MEDIO Then
                
        'Le o Custo do Disponivel
        'Le o ultimo CustoMedio do Mes-Ano aberto
        lErro = CF("CustoMedioAtual_Le", objProduto.sCodigo, dCusto, iFilialEmpresa)
        If lErro <> SUCESSO Then gError 78109
        
        If objEstoqueProdutoSUM.dQuantConserto3 <> 0 Then
            dCustoConserto3 = objEstoqueProdutoSUM.dValorConserto3 / objEstoqueProdutoSUM.dQuantConserto3
        Else
            dCustoConserto3 = 0
        End If
        
        If objEstoqueProdutoSUM.dQuantConsig3 <> 0 Then
            dCustoConsig3 = objEstoqueProdutoSUM.dValorConsig3 / objEstoqueProdutoSUM.dQuantConsig3
        Else
            dCustoConsig3 = 0
        End If
        
        If objEstoqueProdutoSUM.dQuantDemo3 <> 0 Then
            dCustoDemo3 = objEstoqueProdutoSUM.dValorDemo3 / objEstoqueProdutoSUM.dQuantDemo3
        Else
            dCustoDemo3 = 0
        End If
        
        'Outras
        If objEstoqueProdutoSUM.dQuantOutras3 <> 0 Then
            dCustoOutras3 = objEstoqueProdutoSUM.dValorOutras3 / objEstoqueProdutoSUM.dQuantOutras3
        Else
            dCustoOutras3 = 0
        End If
                
        'Beneficiamento
        If objEstoqueProdutoSUM.dQuantBenef3 <> 0 Then
            dCustoBenef3 = objEstoqueProdutoSUM.dValorBenef3 / objEstoqueProdutoSUM.dQuantBenef3
        Else
            dCustoBenef3 = 0
        End If
        
    ElseIf objProduto.iApropriacaoCusto = APROPR_CUSTO_REAL Then
            
        'le o ultimo custo medio producao apurado
        lErro = CF("CustoMedioProducaoApurado_Le", objProduto.sCodigo, dCusto, iFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 69867 And lErro <> 69868 Then gError 78110
                            
        'Le os custos dos Escaninhos
        lErro = CF("CustoMedioProducaoApurado_Le_Escaninhos", objProduto.sCodigo, dCustoConsig, dCustoDemo, dCustoConserto, dCustoOutras, dCustoBenef, dCustoConsig3, dCustoDemo3, dCustoConserto3, dCustoOutras3, dCustoBenef3, iFilialEmpresa)
        If lErro <> SUCESSO Then gError 78111
    
    ElseIf objProduto.iApropriacaoCusto = APROPR_CUSTO_STANDARD Then
    
        'le o custo standard corrente
        lErro = CF("CustoStandard_Le", objProduto.sCodigo, dCusto, iFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 69852 Then gError 78112
           
        'Le o Custo para os Escaninhos
        'Conserto
        If objEstoqueProdutoSUM.dQuantConserto3 <> 0 Then
            dCustoConserto3 = objEstoqueProdutoSUM.dValorConserto3 / objEstoqueProdutoSUM.dQuantConserto3
        Else
            dCustoConserto3 = 0
        End If
            
        'Consignação
        If objEstoqueProdutoSUM.dQuantConsig3 <> 0 Then
            dCustoConsig3 = objEstoqueProdutoSUM.dValorConsig3 / objEstoqueProdutoSUM.dQuantConsig3
        Else
            dCustoConsig3 = 0
        End If
        
        'Demostração
        If objEstoqueProdutoSUM.dQuantDemo3 <> 0 Then
            dCustoDemo3 = objEstoqueProdutoSUM.dValorDemo3 / objEstoqueProdutoSUM.dQuantDemo3
        Else
            dCustoDemo3 = 0
        End If
        
        'Outras
        If objEstoqueProdutoSUM.dQuantOutras3 <> 0 Then
            dCustoOutras3 = objEstoqueProdutoSUM.dValorOutras3 / objEstoqueProdutoSUM.dQuantOutras3
        Else
            dCustoOutras3 = 0
        End If
                
        'Beneficiamento
        If objEstoqueProdutoSUM.dQuantBenef3 <> 0 Then
            dCustoBenef3 = objEstoqueProdutoSUM.dValorBenef3 / objEstoqueProdutoSUM.dQuantBenef3
        Else
            dCustoBenef3 = 0
        End If
    
    End If
    
    Calcula_Valor_Custo = SUCESSO
    
    Exit Function
    
Erro_Calcula_Valor_Custo:
            
    Calcula_Valor_Custo = gErr
    
    Select Case gErr
    
        Case 78108, 78109, 78110, 78111, 78112
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153164)

    End Select

    Exit Function
    
End Function

Private Function Retorna_ValorEstoque_Atual(dValorTotal As Double, sProduto As String, iFilialEmpresa As Integer, iAlmoxarifado As Integer) As Long
'Retorna o valor do Produto no Almoxarifado passado

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim dCusto As Double
Dim dCustoConsig As Double, dCustoDemo As Double, dCustoConserto As Double, dCustoOutras As Double, dCustoBenef As Double, dCustoConsig3 As Double, dCustoDemo3 As Double, dCustoConserto3 As Double, dCustoOutras3 As Double, dCustoBenef3 As Double
Dim dQuantidadeTotal As Double
Dim objEstoqueProdutoSUM As New ClassEstoqueProduto

On Error GoTo Erro_Retorna_ValorEstoque_Atual

    objEstoqueProduto.sProduto = sProduto
    
    'Lê a soma de todas as quantidades para Produto Passado em todos os Almoxarifados da Filial
    lErro = CF("EstoqueProduto_Le_Todos_Almoxarifados_Filial", objEstoqueProdutoSUM, iFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 64014 Then gError 78113
            
    objEstoqueProduto.sProduto = sProduto
    objEstoqueProduto.iAlmoxarifado = iAlmoxarifado
    
    'Lê o Estoque Produto
    lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
    If lErro <> SUCESSO And lErro <> 21306 Then gError 78114
    
    'Calcula o Custo
    lErro = Calcula_Valor_Custo(sProduto, objEstoqueProdutoSUM, dCusto, dCustoConsig, dCustoDemo, dCustoConserto, dCustoOutras, dCustoBenef, dCustoConsig3, dCustoDemo3, dCustoConserto3, dCustoOutras3, dCustoBenef3, iFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 69854 Then gError 78115
            
    'Calcula o Valor Total
    dQuantidadeTotal = objEstoqueProduto.dQuantDisponivel + objEstoqueProduto.dQuantInd + objEstoqueProduto.dQuantRecIndl + objEstoqueProduto.dQuantReservada + objEstoqueProduto.dQuantReservadaConsig + objEstoqueProduto.dQuantDefeituosa
    dValorTotal = (dQuantidadeTotal * dCusto) + (objEstoqueProduto.dQuantConsig3 * dCustoConsig3) + (objEstoqueProduto.dQuantOutras3 * dCustoOutras3) + (objEstoqueProduto.dQuantConserto3 * dCustoConserto3) + (objEstoqueProduto.dQuantDemo3 * dCustoDemo3) + (objEstoqueProduto.dQuantBenef3 * dCustoBenef3)
    
    Retorna_ValorEstoque_Atual = SUCESSO

    Exit Function

Erro_Retorna_ValorEstoque_Atual:

    Retorna_ValorEstoque_Atual = gErr
    
    Select Case gErr

        Case 78113, 78114, 78115
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153165)

    End Select

    Exit Function

End Function

'Tambem está em ctEstoqueProduto tem que ir para rotinas est
Function CustoMedioProducaoApurado_Le_Escaninhos(sProduto As String, dCustoConsig As Double, dCustoDemo As Double, dCustoConserto As Double, dCustoOutras As Double, dCustoBenef As Double, dCustoConsig3 As Double, dCustoDemo3 As Double, dCustoConserto3 As Double, dCustoOutras3 As Double, dCustoBenef3 As Double, iFilialEmpresa As Integer) As Long
'Calcula último Custo Médio de Produção Apurado do Produto passado

Dim lErro As Long
Dim objEstoqueMes As New ClassEstoqueMes
Dim alComando(1 To 10) As Long
Dim iIndice As Integer
Dim objSldMesEst As New ClassSldMesEst

On Error GoTo Erro_CustoMedioProducaoApurado_Le_Escaninhos

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69791
    Next

    objEstoqueMes.iFilialEmpresa = iFilialEmpresa
    
    lErro = CF("EstoqueMes_Le_Apurado", objEstoqueMes)
    If lErro <> SUCESSO And lErro <> 46225 Then gError 69792
    
    'Se não foi apurado nenhum mês ==> descobre o ano em que o estoque foi
    If lErro = 46225 Then
        
        objSldMesEst.sProduto = sProduto
        objSldMesEst.iFilialEmpresa = objEstoqueMes.iFilialEmpresa
        
        'Se não  encontrou mês apurado então lê o SldMesEst para pegar o primeiro Ano
        lErro = CF("SldMesEst_Le_PrimeiroAno", objSldMesEst)
        If lErro <> SUCESSO Then gError 69831
        
        objEstoqueMes.iMes = 0
        objEstoqueMes.iAno = objSldMesEst.iAno
    End If
    
    'calcula o custo medio de produção Benef3 do produto em questão
    lErro = CF("SldMesEst1_Le_CustoMedioBenef3", alComando(1), iFilialEmpresa, objEstoqueMes.iAno, sProduto, objEstoqueMes.iMes, dCustoBenef3)
    If lErro <> SUCESSO Then gError 69793
    
    'calcula o custo medio de produção Conserto3 do produto em questão
    lErro = CF("SldMesEst1_Le_CustoMedioConserto3", alComando(1), iFilialEmpresa, objEstoqueMes.iAno, sProduto, objEstoqueMes.iMes, dCustoConserto3)
    If lErro <> SUCESSO Then gError 69794
    
    'calcula o custo medio de produção Consig3 do produto em questão
    lErro = CF("SldMesEst1_Le_CustoMedioConsig3", alComando(1), iFilialEmpresa, objEstoqueMes.iAno, sProduto, objEstoqueMes.iMes, dCustoConsig3)
    If lErro <> SUCESSO Then gError 69795
    
    'calcula o custo medio de produção Demo3 do produto em questão
    lErro = CF("SldMesEst1_Le_CustoMedioDemo3", alComando(1), iFilialEmpresa, objEstoqueMes.iAno, sProduto, objEstoqueMes.iMes, dCustoDemo3)
    If lErro <> SUCESSO Then gError 69796
    
    'calcula o custo medio de produção Outros3  do produto em questão
    lErro = CF("SldMesEst1_Le_CustoMedioOutros3", alComando(1), iFilialEmpresa, objEstoqueMes.iAno, sProduto, objEstoqueMes.iMes, dCustoOutras3)
    If lErro <> SUCESSO Then gError 69797
        
    'calcula o custo medio de produção Benef do produto em questão
    lErro = CF("SldMesEst2_Le_CustoMedioBenef", alComando(1), iFilialEmpresa, objEstoqueMes.iAno, sProduto, objEstoqueMes.iMes, dCustoBenef)
    If lErro <> SUCESSO Then gError 69798
    
    'calcula o custo medio de produção Conserto  do produto em questão
    lErro = CF("SldMesEst2_Le_CustoMedioConserto", alComando(1), iFilialEmpresa, objEstoqueMes.iAno, sProduto, objEstoqueMes.iMes, dCustoConserto)
    If lErro <> SUCESSO Then gError 69799
    
    'calcula o custo medio de produção Consug  do produto em questão
    lErro = CF("SldMesEst2_Le_CustoMedioConsig", alComando(1), iFilialEmpresa, objEstoqueMes.iAno, sProduto, objEstoqueMes.iMes, dCustoConsig)
    If lErro <> SUCESSO Then gError 69800
    
    'calcula o custo medio de produção Demo  do produto em questão
    lErro = CF("SldMesEst2_Le_CustoMedioDemo", alComando(1), iFilialEmpresa, objEstoqueMes.iAno, sProduto, objEstoqueMes.iMes, dCustoDemo)
    If lErro <> SUCESSO Then gError 69801
    
    'calcula o custo medio de produção Outros do produto em questão
    lErro = CF("SldMesEst2_Le_CustoMedioOutros", alComando(1), iFilialEmpresa, objEstoqueMes.iAno, sProduto, objEstoqueMes.iMes, dCustoOutras)
    If lErro <> SUCESSO Then gError 69802

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    CustoMedioProducaoApurado_Le_Escaninhos = SUCESSO

    Exit Function

Erro_CustoMedioProducaoApurado_Le_Escaninhos:

    CustoMedioProducaoApurado_Le_Escaninhos = gErr

    Select Case gErr
        
        Case 69791
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69792 To 69802 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153166)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function ObterSaldoProdDtAlmox(dQuantidade As Double, sProduto As String, iAlmoxarifado As Integer, dtData As Date, iFilialEmpresa As Integer) As Long
'Para obter o saldo em estoque da propria empresa de um produto em um almoxarifado em uma data

Dim lErro As Long, objRegInventario As New ClassRegInventario

On Error GoTo Erro_ObterSaldoProdDtAlmox

    If (sProdutoAntInv <> sProduto Or iInvAlmoxarifadoAntInv <> iAlmoxarifado Or dtDataAntInv <> dtData) Then
    
        With objRegInventario
            .iFilialEmpresa = iFilialEmpresa
            .sProduto = sProduto
            .iAlmoxarifado = iAlmoxarifado
            .dtData = dtData
        End With
        
        lErro = CF("RegInventario_Obter_Quantidades_Almox", objRegInventario)
        If lErro <> SUCESSO Then gError 81618
        
        sProdutoAntInv = sProduto
        iInvAlmoxarifadoAntInv = iAlmoxarifado
        dtDataAntInv = dtData
        
        objRegInventario.dQtdeDeTercConosco = objRegInventario.dQuantBenef3 + objRegInventario.dQuantConserto3 + objRegInventario.dQuantConsig3 + objRegInventario.dQuantDemo3 + objRegInventario.dQuantOutras3
        objRegInventario.dQtdeNossaEmTerc = objRegInventario.dQuantBenef + objRegInventario.dQuantConserto + objRegInventario.dQuantConsig + objRegInventario.dQuantDemo + objRegInventario.dQuantOutras
        
        dQuantidadeNossaInv = objRegInventario.dQuantidadeUMEstoque
        dQuantidadeEmTerc = objRegInventario.dQtdeNossaEmTerc
        dQuantidadeDeTerc = objRegInventario.dQtdeDeTercConosco
        
    End If
        
    dQuantidade = dQuantidadeNossaInv
        
    ObterSaldoProdDtAlmox = SUCESSO
     
    Exit Function
    
Erro_ObterSaldoProdDtAlmox:

    ObterSaldoProdDtAlmox = gErr
     
    Select Case gErr
          
        Case 81618
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153167)
     
    End Select
     
    Exit Function

End Function

Public Function ObterSaldoProdDtAlmoxEmTerc(dQuantidade As Double, sProduto As String, iAlmoxarifado As Integer, dtData As Date, iFilialEmpresa As Integer) As Long
'Para obter o saldo de estoque em 3os alocada a um produto-almoxarifado em uma data

Dim lErro As Long, objRegInventario As New ClassRegInventario

On Error GoTo Erro_ObterSaldoProdDtAlmoxEmTerc

    If (sProdutoAntInv <> sProduto Or iInvAlmoxarifadoAntInv <> iAlmoxarifado Or dtDataAntInv <> dtData) Then
    
        With objRegInventario
            .iFilialEmpresa = iFilialEmpresa
            .sProduto = sProduto
            .iAlmoxarifado = iAlmoxarifado
            .dtData = dtData
        End With
        
        lErro = CF("RegInventario_Obter_Quantidades_Almox", objRegInventario)
        If lErro <> SUCESSO Then gError 81619
        
        sProdutoAntInv = sProduto
        iInvAlmoxarifadoAntInv = iAlmoxarifado
        dtDataAntInv = dtData
        
        objRegInventario.dQtdeDeTercConosco = objRegInventario.dQuantBenef3 + objRegInventario.dQuantConserto3 + objRegInventario.dQuantConsig3 + objRegInventario.dQuantDemo3 + objRegInventario.dQuantOutras3
        objRegInventario.dQtdeNossaEmTerc = objRegInventario.dQuantBenef + objRegInventario.dQuantConserto + objRegInventario.dQuantConsig + objRegInventario.dQuantDemo + objRegInventario.dQuantOutras
        
        dQuantidadeNossaInv = objRegInventario.dQuantidadeUMEstoque
        dQuantidadeEmTerc = objRegInventario.dQtdeNossaEmTerc
        dQuantidadeDeTerc = objRegInventario.dQtdeDeTercConosco
        
    End If
        
    dQuantidade = dQuantidadeEmTerc
        
    ObterSaldoProdDtAlmoxEmTerc = SUCESSO
     
    Exit Function
    
Erro_ObterSaldoProdDtAlmoxEmTerc:

    ObterSaldoProdDtAlmoxEmTerc = gErr
     
    Select Case gErr
          
        Case 81619
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153168)
     
    End Select
     
    Exit Function

End Function

Public Function ObterSaldoProdDtAlmoxDeTerc(dQuantidade As Double, sProduto As String, iAlmoxarifado As Integer, dtData As Date, iFilialEmpresa As Integer) As Long
'Para obter o saldo de estoque de 3os em um produto-almoxarifado em uma data

Dim lErro As Long, objRegInventario As New ClassRegInventario

On Error GoTo Erro_ObterSaldoProdDtAlmoxDeTerc

    If (sProdutoAntInv <> sProduto Or iInvAlmoxarifadoAntInv <> iAlmoxarifado Or dtDataAntInv <> dtData) Then
    
        With objRegInventario
            .iFilialEmpresa = iFilialEmpresa
            .sProduto = sProduto
            .iAlmoxarifado = iAlmoxarifado
            .dtData = dtData
        End With
        
        lErro = CF("RegInventario_Obter_Quantidades_Almox", objRegInventario)
        If lErro <> SUCESSO Then gError 81620
        
        sProdutoAntInv = sProduto
        iInvAlmoxarifadoAntInv = iAlmoxarifado
        dtDataAntInv = dtData
        
        objRegInventario.dQtdeDeTercConosco = objRegInventario.dQuantBenef3 + objRegInventario.dQuantConserto3 + objRegInventario.dQuantConsig3 + objRegInventario.dQuantDemo3 + objRegInventario.dQuantOutras3
        objRegInventario.dQtdeNossaEmTerc = objRegInventario.dQuantBenef + objRegInventario.dQuantConserto + objRegInventario.dQuantConsig + objRegInventario.dQuantDemo + objRegInventario.dQuantOutras
        
        dQuantidadeNossaInv = objRegInventario.dQuantidadeUMEstoque
        dQuantidadeEmTerc = objRegInventario.dQtdeNossaEmTerc
        dQuantidadeDeTerc = objRegInventario.dQtdeDeTercConosco
            
    End If
    
    dQuantidade = dQuantidadeDeTerc
        
    ObterSaldoProdDtAlmoxDeTerc = SUCESSO
     
    Exit Function
    
Erro_ObterSaldoProdDtAlmoxDeTerc:

    ObterSaldoProdDtAlmoxDeTerc = gErr
     
    Select Case gErr
          
        Case 81620
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153169)
     
    End Select
     
    Exit Function

End Function

Public Function ObterSaldoProdDtAlmox1(dQuantidade As Double, sProduto As String, iAlmoxarifado As Integer, dtDataInicio As Date, iFilialEmpresa As Integer, ByVal dtDataFim As Date, ByVal dQuantMinima As Double) As Long
'Vai pesquisar a quantidade em estoque em dtDataInicio e se tiver mais que dQuantMinima vai pesquisar a quantidade na dtDataFim
'dQuantMinima é a quantidade minima que o produto deve ter na data dtDataInicio para continuar a pesquisa na dtDataFim.

Dim lErro As Long, objRegInventario As New ClassRegInventario

On Error GoTo Erro_ObterSaldoProdDtAlmox1

    If (sProdutoAntInv <> sProduto Or iInvAlmoxarifadoAntInv <> iAlmoxarifado Or dtDataAntInv <> dtDataInicio Or dtDataFimAnt <> dtDataFim Or dQuantMinimaAnt <> dQuantMinima) Then
    
        With objRegInventario
            .iFilialEmpresa = iFilialEmpresa
            .sProduto = sProduto
            .iAlmoxarifado = iAlmoxarifado
            .dtData = dtDataInicio
        End With
        
        lErro = CF("Produto_Obter_Quantidades_Almox", objRegInventario, dtDataFim, dQuantMinima)
        If lErro <> SUCESSO And lErro <> 105165 And lErro <> 105168 Then gError 81618
        
        sProdutoAntInv = sProduto
        iInvAlmoxarifadoAntInv = iAlmoxarifado
        dtDataAntInv = dtDataInicio
        dtDataFimAnt = dtDataFim
        dQuantMinimaAnt = dQuantMinima
        
        If lErro = 105165 Or lErro = 105168 Then
            dQuantidadeNossaInv = -1
        Else
            dQuantidadeNossaInv = objRegInventario.dQuantidadeUMEstoque
        End If
        
    End If
        
    dQuantidade = dQuantidadeNossaInv
        
    ObterSaldoProdDtAlmox1 = SUCESSO
     
    Exit Function
    
Erro_ObterSaldoProdDtAlmox1:

    ObterSaldoProdDtAlmox1 = gErr
     
    Select Case gErr
          
        Case 81618
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153170)
     
    End Select
     
    Exit Function

End Function

