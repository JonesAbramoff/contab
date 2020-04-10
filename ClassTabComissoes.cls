VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTabComissoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarobjTela As Object

'Variáveis para colunas do Grid
Private iGrid_Vendedor_Col As Integer
Private iGrid_PercentualComissao_Col As Integer
Private iGrid_ValorBase_Col As Integer
Private iGrid_ValorComissao_Col As Integer
Private iGrid_PercentualEmissao_Col As Integer
Private iGrid_ValorEmissao_Col As Integer
Private iGrid_PercentualBaixa_Col As Integer
Private iGrid_ValorBaixa_Col As Integer
Private iGrid_DiretoIndireto_Col As Integer

Private gColComissoes As New Collection

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property


Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Function Comissao_Automatica_FilialCli_Exibe(objFilialCliente As ClassFilialCliente) As Long
'Calcula todos os dados de comissão automática e exibe para essa FilialCliente
'Chamada quando está mudando FilialCliente na Combo de Filiais (de Cliente) da tela

Dim objComissaoNF As New ClassComissaoNF
Dim objVendedor As New ClassVendedor
Dim lErro As Long

On Error GoTo Erro_Comissao_Automatica_FilialCli_Exibe

    'Calcula comissao automatica para essa FilialCliente
    lErro = Comissao_Automatica_Calcula(objFilialCliente, objVendedor, objComissaoNF)
    If lErro <> SUCESSO Then Error 25667

    'Se Vendedor está definido,
    If objComissaoNF.iCodVendedor > 0 Then

        'Limpa o Grid e exibe dados na primeira linha
        Call Grid_Limpa(objTela.objGridComissoes)
        'Vendedor
        objTela.Controls("GridComissoes").TextMatrix(1, iGrid_Vendedor_Col) = objVendedor.sNomeReduzido
        'Percentuais e valores de comissão
        objTela.Controls("GridComissoes").TextMatrix(1, iGrid_PercentualComissao_Col) = IIf(objComissaoNF.dPercentual <> 0, Format(objComissaoNF.dPercentual, "Percent"), "")
        objTela.Controls("GridComissoes").TextMatrix(1, iGrid_PercentualEmissao_Col) = Format(objComissaoNF.dPercentualEmissao, "Percent")
        objTela.Controls("GridComissoes").TextMatrix(1, iGrid_PercentualBaixa_Col) = Format(1 - objComissaoNF.dPercentualEmissao, "Percent")
        objTela.Controls("GridComissoes").TextMatrix(1, iGrid_ValorBase_Col) = IIf(objComissaoNF.dValorBase <> 0, Format(objComissaoNF.dValorBase, "Standard"), "")
        objTela.Controls("GridComissoes").TextMatrix(1, iGrid_ValorComissao_Col) = IIf(objComissaoNF.dValor <> 0, Format(objComissaoNF.dValor, "Standard"), "")
        objTela.Controls("GridComissoes").TextMatrix(1, iGrid_ValorEmissao_Col) = IIf(objComissaoNF.dValor <> 0, Format(objComissaoNF.dValorEmissao, "Standard"), "")
        objTela.Controls("GridComissoes").TextMatrix(1, iGrid_ValorBaixa_Col) = IIf(objComissaoNF.dValor <> 0, Format(objComissaoNF.dValor - objComissaoNF.dValorEmissao, "Standard"), "")
        objTela.Controls("GridComissoes").TextMatrix(1, iGrid_DiretoIndireto_Col) = IIf(objComissaoNF.iIndireta <> VENDEDOR_DIRETO, VENDEDOR_INDIRETO_STRING, VENDEDOR_DIRETO_STRING)

        'Acrescenta uma linha no Grid Comissoes
        objTela.objGridComissoes.iLinhasExistentes = 1

        'Tulio-10/05/02
        Call Calcula_Valores

    End If

    Comissao_Automatica_FilialCli_Exibe = SUCESSO

    Exit Function

Erro_Comissao_Automatica_FilialCli_Exibe:

    Comissao_Automatica_FilialCli_Exibe = Err

    Select Case Err

        Case 25667

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153567)

    End Select

    Exit Function

End Function
Public Function Comissao_Automatica_Vendedor_Exibe(iLinha As Integer, objFilialCliente As ClassFilialCliente, objVendedor As ClassVendedor) As Long
'Calcula todos os dados de comissão automática e exibe para esse par Vendedor, FilialCliente
'Chamada quando está mudando Vendedor ou acrescentando Vendedor no grid Comissoes

Dim objComissaoNF As New ClassComissaoNF
Dim lErro As Long

On Error GoTo Erro_Comissao_Automatica_Vendedor_Exibe

    'Calcula comissao automatica para essa FilialCliente
    lErro = Comissao_Automatica_Calcula(objFilialCliente, objVendedor, objComissaoNF)
    If lErro <> SUCESSO Then Error 25668

    'Se Vendedor está definido,
    If objComissaoNF.iCodVendedor > 0 Then

        objTela.Controls("GridComissoes").TextMatrix(iLinha, iGrid_Vendedor_Col) = objVendedor.sNomeReduzido
        objTela.Controls("Vendedor").Text = objVendedor.sNomeReduzido
        
        'Percentuais e valores de comissão
        objTela.Controls("GridComissoes").TextMatrix(iLinha, iGrid_PercentualComissao_Col) = IIf(objComissaoNF.dPercentual <> 0, Format(objComissaoNF.dPercentual, "Percent"), "")
        objTela.Controls("GridComissoes").TextMatrix(iLinha, iGrid_PercentualEmissao_Col) = Format(objComissaoNF.dPercentualEmissao, "Percent")
        objTela.Controls("GridComissoes").TextMatrix(iLinha, iGrid_PercentualBaixa_Col) = Format(1 - objComissaoNF.dPercentualEmissao, "Percent")
        objTela.Controls("GridComissoes").TextMatrix(iLinha, iGrid_ValorBase_Col) = IIf(objComissaoNF.dValorBase <> 0, Format(objComissaoNF.dValorBase, "Standard"), "")
        objTela.Controls("GridComissoes").TextMatrix(iLinha, iGrid_ValorComissao_Col) = IIf(objComissaoNF.dValor <> 0, Format(objComissaoNF.dValor, "Standard"), "")
        objTela.Controls("GridComissoes").TextMatrix(iLinha, iGrid_ValorEmissao_Col) = IIf(objComissaoNF.dValor <> 0, Format(objComissaoNF.dValorEmissao, "Standard"), "")
        objTela.Controls("GridComissoes").TextMatrix(iLinha, iGrid_ValorBaixa_Col) = IIf(objComissaoNF.dValor <> 0, Format(objComissaoNF.dValor - objComissaoNF.dValorEmissao, "Standard"), "")
        objTela.Controls("GridComissoes").TextMatrix(iLinha, iGrid_DiretoIndireto_Col) = IIf(objComissaoNF.iIndireta <> VENDEDOR_DIRETO, VENDEDOR_INDIRETO_STRING, VENDEDOR_DIRETO_STRING)

        'Acrescenta uma linha no Grid se for o caso
        If iLinha - objTela.Controls("GridComissoes").FixedRows = objTela.objGridComissoes.iLinhasExistentes Then
            objTela.objGridComissoes.iLinhasExistentes = objTela.objGridComissoes.iLinhasExistentes + 1
        End If

        'Tulio-10/05/02
        Call Calcula_Valores
        
    End If

    Comissao_Automatica_Vendedor_Exibe = SUCESSO

    Exit Function

Erro_Comissao_Automatica_Vendedor_Exibe:

    Comissao_Automatica_Vendedor_Exibe = Err

    Select Case Err

        Case 25668

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153568)

    End Select

    Exit Function

End Function

'Substitui ValorBase_Calcula
'Substitui ComissaoAutomatica_Obter_NF
Public Function Comissao_Automatica_Calcula(objFilialCliente As ClassFilialCliente, objVendedor As ClassVendedor, objComissaoNF As ClassComissaoNF) As Long
'Calcula valores default de objComissaoNF (vendedor, comissoes, valores)
'Para ser chamada quando muda Filial Cliente ou quando muda Vendedor
'Ela calcula o valor base e valores e chama Comissao_Obter_Vendedor_Percentuais para obter Vendedor e percentuais

Dim dValor As Double
Dim dValorSeguro As Double
Dim dValorFrete As Double
Dim dValorIPI As Double
Dim dValorDespesas As Double
Dim lErro As Long

On Error GoTo Erro_Comissao_Automatica_Calcula

    'Tenta encontrar valores default para Vendedor e percentuais
    lErro = Comissao_Obter_Vendedor_Percentuais(objFilialCliente, objVendedor, objComissaoNF)
    If lErro <> SUCESSO Then Error 25669

    'Encontra valor base, valor comissão e valor emissão a partir dos percentuais
    lErro = Comissao_Obter_Valores(objComissaoNF, objVendedor)
    If lErro <> SUCESSO Then Error 25725

    Comissao_Automatica_Calcula = SUCESSO

    Exit Function

Erro_Comissao_Automatica_Calcula:

    Comissao_Automatica_Calcula = Err

    Select Case Err

        Case 25669, 25725  'tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153569)

    End Select

    Exit Function

End Function
Private Function Comissao_Obter_Valores(objComissaoNF As ClassComissaoNF, objVendedor As ClassVendedor) As Long
'Encontra valor base, valor comissão e valor emissão a partir dos percentuais (passados),
'dos valores de campos da tela e das condições do Vendedor (passado)

Dim dValor As Double
Dim dValorSeguro As Double
Dim dValorFrete As Double
Dim dValorIPI As Double
Dim dValorDespesas As Double
Dim dValorICMSST As Double
Dim lErro As Long
Dim objTribDoc As Object
Dim dValorAux As Double

On Error GoTo Erro_Comissao_Obter_Valores

    'Zera valores
    objComissaoNF.dValorBase = 0
    objComissaoNF.dValor = 0
    objComissaoNF.dValorEmissao = 0

    'Recolhe o Valor total da tela
    dValor = StrParaDbl(objTela.Controls("ValorTotal"))

    'Recolhes os outros valores
    dValorFrete = StrParaDbl(objTela.Controls("ValorFrete"))
    dValorSeguro = StrParaDbl(objTela.Controls("ValorSeguro"))
    dValorIPI = StrParaDbl(objTela.Controls("IPIValor1"))
    dValorDespesas = StrParaDbl(objTela.Controls("ValorDespesas"))
    dValorICMSST = StrParaDbl(objTela.Controls("ICMSSubstValor1"))

    If objVendedor.iCodigo <> 0 Then

        'De acordo com os dados do Vendedor calcula o valor base
        If objVendedor.iComissaoSobreTotal <> 1 Then
            If objVendedor.iComissaoSeguro = 0 Then dValor = dValor - dValorSeguro
            If objVendedor.iComissaoFrete = 0 Then dValor = dValor - dValorFrete
            If objVendedor.iComissaoIPI = 0 Then dValor = dValor - dValorIPI
            If objVendedor.iComissaoICM = 0 Then dValor = dValor - dValorDespesas
            If gobjCRFAT.iComisRetiraCreditosImp = MARCADO Then
                lErro = Obtem_Tributacao_Tela(objTribDoc)
                If lErro = SUCESSO Then
                    dValor = dValor - objTribDoc.dPISValor - objTribDoc.dCOFINSValor - objTribDoc.dICMSCredito - objTribDoc.dICMSInterestVlrFCPUFDest - objTribDoc.dICMSInterestVlrUFDest - objTribDoc.dICMSInterestVlrUFRemet
                End If
            End If
            dValor = dValor - dValorICMSST
        End If
        Call Obtem_ValorItens_NaoFat(dValorAux)
        'Abate os valores que não são considerados para o financeiro
        dValor = dValor - dValorAux

        objComissaoNF.dValorBase = dValor
        objComissaoNF.dValor = Round(objComissaoNF.dValorBase * objComissaoNF.dPercentual, 2)
        objComissaoNF.dValorEmissao = Round(objComissaoNF.dValorBase * objComissaoNF.dPercentual * objComissaoNF.dPercentualEmissao, 2)

    End If

    Comissao_Obter_Valores = SUCESSO

    Exit Function

Erro_Comissao_Obter_Valores:

    Comissao_Obter_Valores = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153570)

    End Select

    Exit Function

End Function

'SUBSTITUI ComissaoAutomatica_Obter_Info2
Public Function Comissao_Obter_Vendedor_Percentuais(objFilialCliente As ClassFilialCliente, objVendedor As ClassVendedor, objComissaoNF As ClassComissaoNF) As Long
'retorna em objVendedor e objComissao valores de Vendedor e Percentuais default para geracao de comissao automatica à partir de Vendedor e FilialCliente informados
'se objVendedor.iCodigo = 0 entao obtém o vendedor padrao da filial cliente ou do tipo cliente

Dim lErro As Long
Dim objTipoCliente As New ClassTipoCliente
Dim objCliente As New ClassCliente

On Error GoTo Erro_Comissao_Obter_Vendedor_Percentuais

    objComissaoNF.iCodVendedor = objVendedor.iCodigo
    objComissaoNF.dPercentual = 0
    objComissaoNF.dPercentualEmissao = 0

    If objFilialCliente.lCodCliente <> 0 And objFilialCliente.iCodFilial <> 0 Then

        'se a filial cliente tem o vendedor e o % da comissao
        If objFilialCliente.iVendedor <> 0 And objFilialCliente.dComissaoVendas <> 0 Then

            'se o vendedor ainda nao foi selecionado
            If objComissaoNF.iCodVendedor = 0 Then objComissaoNF.iCodVendedor = objFilialCliente.iVendedor

            objComissaoNF.dPercentual = objFilialCliente.dComissaoVendas

        Else 'a filial cliente nao tem tudo vou ter que tentar o tipo de cliente

            If objComissaoNF.iCodVendedor = 0 Then objComissaoNF.iCodVendedor = objFilialCliente.iVendedor
            objComissaoNF.dPercentual = objFilialCliente.dComissaoVendas

            objCliente.lCodigo = objFilialCliente.lCodCliente

            'Lê o Cliente
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO And lErro <> 12293 Then Error 25670

            'Não achou o Cliente --> erro
            If lErro = 12293 Then Error 25671

            If objCliente.iTipo <> 0 Then

                objTipoCliente.iCodigo = objCliente.iTipo

                'Lê o Tipo de Cliente
                lErro = CF("TipoCliente_Le", objTipoCliente)
                If lErro <> SUCESSO And lErro <> 19062 Then Error 25672

                'Não achou o Tipo de Cliente --> erro
                If lErro = 19062 Then Error 25673

                If objComissaoNF.iCodVendedor = 0 Then objComissaoNF.iCodVendedor = objTipoCliente.iVendedor
                If objComissaoNF.dPercentual = 0 Then objComissaoNF.dPercentual = objTipoCliente.dComissaoVendas

            End If

        End If

    End If

    'se o vendedor nao veio já carregado e foi obtido da filial do cliente ou do tipo de cliente
    If objVendedor.iCodigo = 0 And objComissaoNF.iCodVendedor <> 0 Then

        objVendedor.iCodigo = objComissaoNF.iCodVendedor

        'Lê o Vendedor
        lErro = CF("Vendedor_Le", objVendedor)
        If lErro <> SUCESSO And lErro <> 12582 Then Error 25674

        'Não achou o Vendedor --> erro
        If lErro = 12582 Then Error 25675

    End If

    'Verifica se achou o Vendedor
    If objVendedor.iCodigo <> 0 Then

        'Se ainda nao achou o % de comissao usar o do Vendedor
        If objComissaoNF.dPercentual = 0 Then objComissaoNF.dPercentual = objVendedor.dPercComissao

        lErro = CF("Vendedor_Comissao_Customiza1", objVendedor, objComissaoNF, objFilialCliente)
        If lErro <> SUCESSO Then Error 20866

        objComissaoNF.dPercentualEmissao = objVendedor.dPercComissaoEmissao

    End If

    Comissao_Obter_Vendedor_Percentuais = SUCESSO

    Exit Function

Erro_Comissao_Obter_Vendedor_Percentuais:

    Comissao_Obter_Vendedor_Percentuais = Err

    Select Case Err

        Case 20866, 25670, 25672, 25674

        Case 25671
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objCliente.lCodigo)

        Case 25673
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_CLIENTE_NAO_CADASTRADO", Err, objTipoCliente.iCodigo)

        Case 25675
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", Err, objVendedor.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153571)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_GridComissoes(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridComissoes

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_Vendedor_Col
            lErro = Saida_Celula_Vendedor(objGridInt)
            If lErro <> SUCESSO Then gError 25680

        Case iGrid_ValorBase_Col
            lErro = Saida_Celula_ValorBase(objGridInt)
            If lErro <> SUCESSO Then gError 25681

        Case iGrid_PercentualComissao_Col
            lErro = Saida_Celula_PercentualComissao(objGridInt)
            If lErro <> SUCESSO Then gError 25682

        Case iGrid_PercentualEmissao_Col
            lErro = Saida_Celula_PercentualEmissao(objGridInt)
            If lErro <> SUCESSO Then gError 25683

        Case iGrid_ValorComissao_Col
            lErro = Saida_Celula_ValorComissao(objGridInt)
            If lErro <> SUCESSO Then gError 25684

        Case iGrid_ValorEmissao_Col
            lErro = Saida_Celula_ValorEmissao(objGridInt)
            If lErro <> SUCESSO Then gError 25685

        Case iGrid_DiretoIndireto_Col
            lErro = Saida_Celula_DiretoIndireto(objGridInt)
            If lErro <> SUCESSO Then gError 98987

    End Select

    Saida_Celula_GridComissoes = SUCESSO

    Exit Function

Erro_Saida_Celula_GridComissoes:

    Saida_Celula_GridComissoes = gErr

    Select Case gErr

        Case 25680, 25681, 25682, 25683, 25684, 25685, 98987
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153572)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Vendedor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Vendedor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim dValorBase As Double
Dim dValorComissao As Double

On Error GoTo Erro_Saida_Celula_Vendedor

    Set objGridInt.objControle = objTela.Controls("Vendedor")

    'Verifica se vendedor está preenchido
    If Len(Trim(objTela.Controls("Vendedor").Text)) > 0 Then

        'Verifica se Vendedor existe
        lErro = TP_Vendedor_Grid(objTela.Controls("Vendedor"), objVendedor)
        If lErro <> SUCESSO And lErro <> 25018 And lErro <> 25020 Then Error 25686

        If lErro = 25018 Then Error 25687
        If lErro = 25020 Then Error 25688

        'Se houve mudança de Vendedor em relação ao que tinha no Grid
        If Trim(objTela.Controls("Vendedor").Text) <> Trim(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_Vendedor_Col)) Then

            'Preenche a linha de Vendedor com dados default
            lErro = VendedorLinha_Preenche(objVendedor)
            If lErro <> SUCESSO Then Error 25689

            objTela.Controls("ComissaoAutomatica").Value = vbUnchecked


        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 25690

    Saida_Celula_Vendedor = SUCESSO

    Exit Function

Erro_Saida_Celula_Vendedor:

    Saida_Celula_Vendedor = Err

    Select Case Err

        Case 25686, 25689, 25690
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 25687 'Não encontrou nome reduzido de vendedor no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR")

            If vbMsgRes = vbYes Then

                'Preenche objVendedor com nome reduzido
                objVendedor.sNomeReduzido = objTela.Controls("Vendedor").Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a tela de Vendedores
                Call Chama_Tela("Vendedores", objVendedor)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 25688 'Não encontrou codigo do vendedor no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR")

            If vbMsgRes = vbYes Then

                'Prenche objVendedor com codigo
                objVendedor.iCodigo = CDbl(objTela.Controls("Vendedor").Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a tela de Vendedores
                Call Chama_Tela("Vendedores", objVendedor)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153573)

    End Select

    Exit Function

End Function

Public Function VendedorLinha_Preenche(objVendedor As ClassVendedor) As Long
'preenche a linha corrente do grid de comissoes p/o vendedor passado com valores default

Dim lErro As Long, iIndice As Integer
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_VendedorLinha_Preenche

    'chama rotina que ira verificar se o vendedor ja esta no grid
    lErro = VendedorLinha_Preenche1(objVendedor)
    If lErro <> SUCESSO Then gError 98989
    
    objFilialCliente.iCodFilial = Codigo_Extrai(objTela.Controls("Filial"))

    If Len(Trim(objTela.Controls("Cliente"))) > 0 And objFilialCliente.iCodFilial <> 0 Then
        'Lê a FilialCliente
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Trim(objTela.Controls("Cliente")), objFilialCliente)
        If lErro <> SUCESSO Then Error 25692

    End If

    'Exibe comissão automática vinculada a FilialCliente / Vendedor
    lErro = Comissao_Automatica_Vendedor_Exibe(objTela.Controls("GridComissoes").Row, objFilialCliente, objVendedor)
    If lErro <> SUCESSO Then Error 25693

    'Tulio-10/05/02
    Call Calcula_Valores

    VendedorLinha_Preenche = SUCESSO

    Exit Function

Erro_VendedorLinha_Preenche:

    VendedorLinha_Preenche = gErr

    Select Case gErr

        Case 25692, 25693, 98989 'tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153574)

    End Select

    Exit Function

End Function

Public Function VendedorLinha_Preenche1(objVendedor As ClassVendedor) As Long
'Verifica se o vendedor se repete com a flag indireto
'objVendedor eh um parametro de INPUT que traz os

Dim bExisteVendedor As Boolean
Dim iIndice As Integer

On Error GoTo Erro_VendedorLinha_Preenche1

    'Loop no GridComissoes
    For iIndice = 1 To objTela.objGridComissoes.iLinhasExistentes

        'Se o vendedor comparece em outra linha
        If objTela.Controls("GridComissoes").Row <> iIndice And UCase(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_Vendedor_Col)) = UCase(objVendedor.sNomeReduzido) Then

            gError 101636

'O código abaixo foi comentado, pois foi decidido que o vendedor não poderá se repetir no grid,
'entretanto ele não foi apagado, pois caso seja decidido que o grid deve permitir o mesmo vendedor,
'desde que o tipo de comissão seja diferente, basta descomentar o trecho abaixo
'            'achou vendedor
'            bExisteVendedor = True
'
'            'se o campo diretoIndireto da linha atual estiver preenchido com o mesmo da linha atual do loop
'            If Len(Trim(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_DiretoIndireto_Col))) > 0 And objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_DiretoIndireto_Col) = objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_DiretoIndireto_Col) Then gError 25691
'
'            'senao, o campo direto/indireto da linha atual recebe o q nao esta preenchido no campo da linha iindice
'            If objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_DiretoIndireto_Col) = objTela.Controls("DiretoIndireto").List(objTela.Controls("DiretoIndireto").ItemData(VENDEDOR_DIRETO)) Then
'
'                objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_DiretoIndireto_Col) = objTela.Controls("DiretoIndireto").List(objTela.Controls("DiretoIndireto").ItemData(VENDEDOR_INDIRETO))
'
'            Else
'
'                objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_DiretoIndireto_Col) = objTela.Controls("DiretoIndireto").List(objTela.Controls("DiretoIndireto").ItemData(VENDEDOR_DIRETO))
'
'            End If
            
        End If
    
    Next

    'Esse trecho tb não foi apagado pelo mesmo motivo já citado acima
    'se nao encontrou o vendedor em outra linha e o campo direto/indireto esta em branco, seta o campo como direto
    'If bExisteVendedor = False And Len(Trim(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_DiretoIndireto_Col))) = 0 Then objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_DiretoIndireto_Col) = objTela.Controls("DiretoIndireto").List(objTela.Controls("DiretoIndireto").ItemData(VENDEDOR_DIRETO))
       
    VendedorLinha_Preenche1 = SUCESSO

    Exit Function

Erro_VendedorLinha_Preenche1:

    VendedorLinha_Preenche1 = gErr

    Select Case gErr

        Case 101636
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_JA_EXISTENTE", gErr, objVendedor.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153575)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorBase(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor Base do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentual As Double
Dim dValorBase As Double
Dim dValorComissao As Double
Dim lTamanho As Long
Dim dValorGrid As Double
Dim dValorCampo As Double

On Error GoTo Erro_Saida_Celula_ValorBase



    Set objGridInt.objControle = objTela.Controls("ValorBase")

    dValorCampo = StrParaDbl(objTela.Controls("ValorBase").Text)
    dValorGrid = StrParaDbl(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorBase_Col))

    'Se houve mudança de ValorBase em relação ao que tinha no Grid
    If dValorCampo <> dValorGrid Then
        objTela.Controls("ComissaoAutomatica").Value = vbUnchecked
    End If

    'Verifica se valor base está preenchido
    If Len(objTela.Controls("ValorBase").ClipText) > 0 Then

        'Critica se valor base é positivo
        lErro = Valor_NaoNegativo_Critica(objTela.Controls("ValorBase").Text)
        If lErro <> SUCESSO Then Error 25696

        dValorBase = CDbl(objTela.Controls("ValorBase").Text)

        'Mostra na tela o ValorBase formatado
        objTela.Controls("ValorBase").Text = Format(dValorBase, "Standard")

        lTamanho = Len(Trim(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_PercentualComissao_Col)))

        'Verifica se percentual comissao está preenchido
        If lTamanho > 0 Then

            'Percentual de comissao
            dPercentual = PercentParaDbl(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_PercentualComissao_Col))

            'Calcula o valor da comissao
            dValorComissao = dPercentual * dValorBase

            'Mostra na tela o valor da comissao
            objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorComissao_Col) = Format(dValorComissao, "Standard")

        ElseIf Len(Trim(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorComissao_Col))) > 0 Then

            dValorComissao = CDbl(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorComissao_Col))

            If dValorBase < dValorComissao Then Error 25697

            'Calcula percentual de comissao
            If dValorBase > 0 Then dPercentual = dValorComissao / dValorBase

            'Mostra o percentual da comissao na tela
            objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_PercentualComissao_Col) = Format(dPercentual, "Percent")

        End If

        'Calcula valores (percentuais) de emissao e baixa
        If Len(Trim(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorComissao_Col))) > 0 Then
            lErro = EmissaoBaixa_Calcula(dValorComissao)
            If lErro <> SUCESSO Then Error 25698
        End If

        'Acrescenta uma linha no Grid se for o caso
        If objTela.Controls("GridComissoes").Row - objTela.Controls("GridComissoes").FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    ElseIf Len(objTela.Controls("ValorBase").ClipText) = 0 Then

        lErro = Limpa_Valores()
        If lErro <> SUCESSO Then Error 25699

    End If

    If PercentParaDbl(objTela.Controls("TotalPercentualComissao").Caption) > 1 Then Error 25703

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 25702

    'Tulio-10/05/02
    Call Calcula_Valores
    
    Saida_Celula_ValorBase = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorBase:

    Saida_Celula_ValorBase = Err

    Select Case Err

        Case 25696, 25698, 25699, 25702   'tratado na rotina chamada
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 25697
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_COMISSAO_MAIOR_VALORBASE", Err, dValorComissao, dValorBase)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 25703
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TOTAL_PERCENTUAIS_MAIOR_100", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153576)

    End Select

    Exit Function

End Function

Private Function EmissaoBaixa_Calcula(dValorComissao As Double) As Long

Dim lTamanho As Long
Dim dPercentualEm As Double
Dim dValorEmissao As Double
Dim lErro As Long

On Error GoTo Erro_EmissaoBaixa_Calcula

    lTamanho = Len(Trim(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_PercentualEmissao_Col)))
    'Se o percentual de emissão estiver preenchido
    If lTamanho > 0 Then
        'Calcula valores da emissão e da baixa
        dPercentualEm = PercentParaDbl(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_PercentualEmissao_Col))
        dValorEmissao = dPercentualEm * dValorComissao
        objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorEmissao_Col) = Format(dValorEmissao, "Standard")
        objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorBaixa_Col) = Format(dValorComissao - dValorEmissao, "Standard")
    'Se o valor na emissão estiver preenchido
    ElseIf Len(Trim(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorEmissao_Col))) > 0 Then
        'Calcula o Percentual da emissão e o percentual e o valor da baixa
        dValorEmissao = CDbl(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorEmissao_Col))

        If dValorEmissao > dValorComissao Then Error 25704

        objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorBaixa_Col) = Format(dValorComissao - dValorEmissao, "Standard")

        If dValorComissao > 0 Then dPercentualEm = dValorEmissao / dValorComissao
        
        objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_PercentualEmissao_Col) = Format(dPercentualEm, "Percent")
        objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_PercentualBaixa_Col) = Format(1 - dPercentualEm, "Percent")

    End If

    EmissaoBaixa_Calcula = SUCESSO

    Exit Function

Erro_EmissaoBaixa_Calcula:

    EmissaoBaixa_Calcula = Err

    Select Case Err

        Case 25704
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_COMISSAO_EMISSAO_MAIOR", Err, dValorEmissao, dValorComissao)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153577)

    End Select

    Exit Function

End Function

Private Function Limpa_Valores() As Long

Dim iIndice As Integer
Dim dSomaPercentual As Double
Dim lComprimento As Long
Dim lErro As Long

On Error GoTo Erro_Limpa_Valores

    objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorComissao_Col) = ""
    objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorEmissao_Col) = ""
    objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorBaixa_Col) = ""

    Limpa_Valores = SUCESSO

    Exit Function

Erro_Limpa_Valores:

    Limpa_Valores = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153578)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PercentualComissao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual Comissões do grid que está deixando de ser o corrente

Dim lErro As Long
Dim dPercentual As Double
Dim dValorBase As Double
Dim dValorComissao As Double
Dim dPercentualCampo As Double
Dim dPercentualGrid As Double


On Error GoTo Erro_Saida_Celula_PercentualComissao

    Set objGridInt.objControle = objTela.Controls("PercentualComissao")


    dPercentualGrid = PercentParaDbl(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_PercentualComissao_Col))
    dPercentualGrid = dPercentualGrid * 100

    dPercentualCampo = StrParaDbl(objTela.Controls("PercentualComissao").Text)
    
    'Se houve mudança de Percentual em relação ao que tinha no Grid
    If dPercentualGrid <> dPercentualCampo Then
        objTela.Controls("ComissaoAutomatica").Value = vbUnchecked
    End If

    'Verifica se o percentual está preenchido
    If Len(objTela.Controls("PercentualComissao").ClipText) > 0 Then

        'Critica se é porcentagem
        lErro = Porcentagem_Critica(objTela.Controls("PercentualComissao").Text)
        If lErro <> SUCESSO Then Error 25705

        dPercentual = CDbl(objTela.Controls("PercentualComissao").Text)
        
        'Mostra na tela o percentual formatado
        objTela.Controls("PercentualComissao").Text = Format(dPercentual, "Fixed")

        'Verifica se ValorBase correspondente está preenchido
        If Len(Trim(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorBase_Col))) > 0 Then

            dValorBase = CDbl(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorBase_Col))

            'Calcula o ValorComissao
            dValorComissao = dPercentual * dValorBase / 100

            'Coloca o ValorComissao na tela
            objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorComissao_Col) = Format(dValorComissao, "Standard")

            'Coloca na tela valores (percentuais) de comissao na emissao e na baixa
            lErro = EmissaoBaixa_Calcula(dValorComissao)
            If lErro <> SUCESSO Then Error 25706

        End If

        'Acrescenta uma linha no Grid se for o caso
        If objTela.Controls("GridComissoes").Row - objTela.Controls("GridComissoes").FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    ElseIf Len(objTela.Controls("PercentualComissao").ClipText) = 0 Then

        lErro = Limpa_Valores()
        If lErro <> SUCESSO Then Error 25707

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 25708

    'Tulio-10/05/02
    Call Calcula_Valores
    
    Saida_Celula_PercentualComissao = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentualComissao:

    Saida_Celula_PercentualComissao = Err

    Select Case Err

        Case 25705, 25706, 25707, 25708 'tratado na rotina chamada
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 25709
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TOTAL_PERCENTUAIS_MAIOR_100", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153579)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorComissao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentual As Double
Dim dValorBase As Double
Dim dValorComissao As Double
Dim lComprimento As Long
Dim dValorGrid As Double
Dim dValorCampo As Double

On Error GoTo Erro_Saida_Celula_ValorComissao

    Set objGridInt.objControle = objTela.Controls("ValorComissao")

    dValorCampo = StrParaDbl(objTela.Controls("ValorComissao").Text)
    dValorGrid = StrParaDbl(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorComissao_Col))

    'Se houve mudança de ValorBase em relação ao que tinha no Grid
    If dValorCampo <> dValorGrid Then
        objTela.Controls("ComissaoAutomatica").Value = vbUnchecked
    End If


    'Verifica se valor está preenchido
    If Len(objTela.Controls("ValorComissao").ClipText) > 0 Then

        'Critica se valor é positivo
        lErro = Valor_NaoNegativo_Critica(objTela.Controls("ValorComissao").Text)
        If lErro <> SUCESSO Then Error 25711

        dValorComissao = CDbl(objTela.Controls("ValorComissao").Text)

        'Mostra na tela o Valor
        objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorComissao_Col) = Format(dValorComissao, "Fixed")

        'Mostra valores (percentuais) de comissoes na emissao / baixa
        lErro = EmissaoBaixa_Calcula(dValorComissao)
        If lErro <> SUCESSO Then Error 25712

        'Acrescenta uma linha no Grid se for o caso
        If objTela.Controls("GridComissoes").Row - objTela.Controls("GridComissoes").FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    ElseIf Len(objTela.Controls("ValorComissao").ClipText) = 0 Then

        lErro = Limpa_Valores()
        If lErro <> SUCESSO Then Error 25713

    End If

    'Verifica se valor base correspondente está preenchido
    If Len(Trim(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorBase_Col))) > 0 Then

        dValorBase = CDbl(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorBase_Col))

        If dValorBase < dValorComissao Then Error 25714

        'Calcula percent de comissao
        If dValorBase > 0 Then dPercentual = (dValorComissao / dValorBase)

        'Mostra o percentual da comissao na tela
        objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_PercentualComissao_Col) = Format(dPercentual, "Percent")

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 25715

    'Tulio-10/05/02
    Call Calcula_Valores
    
    Saida_Celula_ValorComissao = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorComissao:

    Saida_Celula_ValorComissao = Err

    Select Case Err

        Case 25711, 25712, 25713, 25715
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 25714
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_COMISSAO_MAIOR_VALORBASE", Err, dValorComissao, dValorBase)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 25716
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TOTAL_PERCENTUAIS_MAIOR_100", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153580)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PercentualEmissao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual Comissões do grid que está deixando de ser o corrente

Dim lErro As Long
Dim dPercentual As Double
Dim dValorComissao As Double
Dim dPercentualCampo As Double
Dim dPercentualGrid As Double

On Error GoTo Erro_Saida_Celula_PercentualEmissao

    Set objGridInt.objControle = objTela.Controls("PercentualEmissao")


    dPercentualGrid = PercentParaDbl(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_PercentualEmissao_Col))
    dPercentualGrid = dPercentualGrid * 100

    dPercentualCampo = StrParaDbl(objTela.Controls("PercentualEmissao").Text)

    'Se houve mudança de Percentual em relação ao que tinha no Grid
    If dPercentualGrid <> dPercentualCampo Then
        objTela.Controls("ComissaoAutomatica").Value = vbUnchecked
    End If

    If Len(objTela.Controls("PercentualEmissao").ClipText) > 0 Then
    
        'Critica se é porcentagem
        lErro = Porcentagem_Critica(objTela.Controls("PercentualEmissao").Text)
        If lErro <> SUCESSO Then Error 25718
    
        dPercentual = CDbl(objTela.Controls("PercentualEmissao").Text)
    
        'Mostra na tela o percentual formatado
        objTela.Controls("PercentualEmissao").Text = Format(dPercentual, "Fixed")
    
        dPercentual = CDbl(objTela.Controls("PercentualEmissao").Text) / 100
    
        'Mostra percentual de baixa
        objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_PercentualBaixa_Col) = Format(1 - dPercentual, "Percent")
    
        'Executa o abandona célula
        lErro = Grid_Abandona_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 25719
    
        'Se Valor Comissão está preenchido,
        If Len(Trim(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorComissao_Col))) > 0 Then
    
            dValorComissao = CDbl(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorComissao_Col))
    
            'Mostra valores de comissoes na emissao/baixa
            lErro = EmissaoBaixa_Calcula(dValorComissao)
            If lErro <> SUCESSO Then Error 25720
    
        End If
        
    Else
    
        'Executa o abandona célula
        lErro = Grid_Abandona_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 25731
        
    End If

    Saida_Celula_PercentualEmissao = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentualEmissao:

    Saida_Celula_PercentualEmissao = Err

    Select Case Err

        Case 25718, 25719, 25720, 25731
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153581)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorEmissao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor Emissão do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dPercEmissao As Double
Dim dValorEmissao As Double
Dim dValorComissao As Double

On Error GoTo Erro_Saida_Celula_ValorEmissao

    Set objGridInt.objControle = objTela.Controls("ValorEmissao")

    If Len(Trim(objTela.Controls("ValorEmissao").ClipText)) <> 0 Then

        'Critica se valor emissão é não-negativo
        lErro = Valor_NaoNegativo_Critica(objTela.Controls("ValorEmissao").Text)
        If lErro <> SUCESSO Then Error 25721
    
        dValorEmissao = CDbl(objTela.Controls("ValorEmissao").Text)
    
        'Mostra na tela o Valor
        objTela.Controls("ValorEmissao").Text = Format(dValorEmissao, "Fixed")
    
        dValorEmissao = CDbl(objTela.Controls("ValorEmissao").Text)
    
        'Verifica se valor Comissao correspondente está preenchido
        If Len(Trim(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorComissao_Col))) > 0 Then
    
            dValorComissao = CDbl(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorComissao_Col))
    
            If dValorEmissao > dValorComissao Then Error 25722
    
            'Mostra valor de comissao na baixa
            objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_ValorBaixa_Col) = Format(dValorComissao - dValorEmissao, "Standard")
    
            'Mostra percentuais de comissao na emissao e na baixa
            If dValorComissao > 0 Then
            
                dPercEmissao = (dValorEmissao / dValorComissao)
                objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_PercentualEmissao_Col) = Format(dPercEmissao, "Percent")
                objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_PercentualBaixa_Col) = Format(1 - dPercEmissao, "Percent")
                
            End If
    
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 25723

    Saida_Celula_ValorEmissao = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorEmissao:

    Saida_Celula_ValorEmissao = Err

    Select Case Err

        Case 25721, 25723
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 25722
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_COMISSAO_EMISSAO_MAIOR", Err, dValorEmissao, dValorComissao)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153582)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DiretoIndireto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula DiretoIndireto do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DiretoIndireto

    'seta o objControle como o controle 'diretoindireto'
    Set objGridInt.objControle = objTela.Controls("DiretoIndireto")
    
    If objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_DiretoIndireto_Col) <> objTela.Controls("DiretoIndireto").Text Then
        objTela.Controls("ComissaoAutomatica").Value = vbUnchecked
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 98988

    'Tulio-10/05/02
    Call Calcula_Valores
    
    Saida_Celula_DiretoIndireto = SUCESSO

    Exit Function

Erro_Saida_Celula_DiretoIndireto:

    Saida_Celula_DiretoIndireto = gErr

    Select Case gErr

        Case 98988
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153583)

    End Select

    Exit Function

End Function

Public Function Valida_Grid_Comissoes() As Long
'Valida os dados do Grid de Comissões

Dim lErro As Long
Dim iIndice As Integer
Dim dTotalValorComissao As Double
Dim dValorTotal As Double
Dim dValorIRRF As Double

On Error GoTo Erro_Valida_Grid_Comissoes

    If objTela.objGridComissoes.iLinhasExistentes > 0 Then

        For iIndice = 1 To objTela.objGridComissoes.iLinhasExistentes

            If Len(Trim(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_Vendedor_Col))) = 0 Then Error 42387

            If Len(Trim(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorComissao_Col))) = 0 Then Error 42388
            
            If Len(Trim(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorEmissao_Col))) = 0 Then Error 25729
            If Len(Trim(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_PercentualEmissao_Col))) = 0 Then Error 25730

        Next

        dTotalValorComissao = StrParaDbl(objTela.Controls("TotalValorComissao").Caption)
        dValorTotal = StrParaDbl(objTela.Controls("ValorTotal").Caption)

        If dTotalValorComissao > dValorTotal Then Error 42389

    End If

    Valida_Grid_Comissoes = SUCESSO

    Exit Function

Erro_Valida_Grid_Comissoes:

    Valida_Grid_Comissoes = Err

    Select Case Err

        Case 25729
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_EMISSAO_GRID_NAO_PREENCHIDO", Err, iIndice)
            
        Case 25730
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTAGEM_EMISSAO_GRID_NAO_PREENCHIDA", Err, iIndice)
        
        Case 42387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_COMISSAO_NAO_INFORMADO", Err, iIndice)

        Case 42388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_COMISSAO_GRID_NAO_PREENCHIDO", Err, iIndice)

        Case 42389
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_TOTAL_COMISSAO_INVALIDO", Err, dTotalValorComissao, dValorTotal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153584)

    End Select

    Exit Function

End Function


Public Function Carrega_Tab_Comissoes(objObjetoTela As Object) As Long
'Coloca os dados do tab de Comissões na tela
'O Objeto passado como INPUT deve possuir uma colecao chamada: colComissoes
'ja carregada para o preenchimento do grid
'O intuito dessa funcao eh marcar/desmarcar a check ComissaoAutomatica

Dim lErro As Long
Dim sClasseTela As String
Dim colComissoesAux As Collection

On Error GoTo Erro_Carrega_Tab_Comissoes

    'Guarda o tipo da classe principal da tela que instanciou TabComissoes
    sClasseTela = TypeName(objObjetoTela)
    
    'Seleciona o tipo da classe
    Select Case sClasseTela
    
        'Se for ClassNFiscal
        Case NOME_CLASSE_NOTAFISCAL
    
            'instancia  com a coleção de comissoes de ClassNFiscal
            Set colComissoesAux = objObjetoTela.ColComissoesNF
        
        'Se for ClassPedidoDeVenda
        Case NOME_CLASSE_PEDIDOVENDA
    
            'Marca a check da tela (só ClassPedidoDeVenda tem a property iComissaoAutomatica)
            objTela.Controls("ComissaoAutomatica").Value = objObjetoTela.iComissaoAutomatica
        
            'instancia  com a coleção de comissoes de ClassPedidoDeVenda
            Set colComissoesAux = objObjetoTela.colComissoes
    
    End Select
        
    Set gColComissoes = colComissoesAux
    
    'Chama a funcao responsavel por carregar o grid
    lErro = Carrega_Grid_Comissoes(colComissoesAux)
    If lErro <> SUCESSO Then gError 98982
    
    Carrega_Tab_Comissoes = SUCESSO

    Exit Function

Erro_Carrega_Tab_Comissoes:

    Carrega_Tab_Comissoes = gErr
    
    Select Case gErr
    
        Case 98982
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153585)
            
    End Select
    
    Exit Function
    
End Function

Public Function Carrega_Grid_Comissoes(colComissoes As Collection) As Long
'Preenche o Grid com as Comissões carregadas em colComissoes
'colComissoes eh um parametro de INPUT

Dim lErro As Long
Dim iIndice As Integer
Dim objComissaoGenerica As Object
Dim objVendedor As New ClassVendedor
Dim sDiretoIndireto As String

On Error GoTo Erro_Carrega_Grid_Comissoes

    'Evita que Comissão gerada automaticamente apareça
    Call Grid_Limpa(objTela.objGridComissoes)

    'Para cada tipo de comissao existente na colecao
    'obs. Provavelmente, so existirao tipos iguais dentro da colecao... independente disso,
    'caso esse tipo possua as mesmas propriedades dos tipos de comissoes existentes no
    'sistema, esse loop se torna correto. A variavel se chama ComissaoGenerica pois nao
    'se sabe exatamente o tipo de comissao que sera manipulado... no entanto, como ja foi dito
    'as propriedades desse tipo ja sao conhecidas...
    For Each objComissaoGenerica In colComissoes

        'Incrementa o indice (q significa a linha corrente)
        iIndice = iIndice + 1

        'guarda o codigo do vendedor no objvendedor para uma posterior leitura
        objVendedor.iCodigo = objComissaoGenerica.iCodVendedor
        
        'eh feita a leitura para validacao do vendedor associado a comissao
        lErro = CF("Vendedor_Le", objVendedor)
        If lErro <> SUCESSO And lErro <> 12582 Then gError 98983

        'Se não encontrou o Vendedor --> erro, possivel problema de integridade...
        'If lErro = 12582 Then gError 98984

        'Atribuicoes para o preenchimento do grid...
        objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_Vendedor_Col) = objVendedor.sNomeReduzido
        objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorBase_Col) = Format(objComissaoGenerica.dValorBase, "Standard")
        objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_PercentualComissao_Col) = Format(objComissaoGenerica.dPercentual, "Percent")
        objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorComissao_Col) = Format(objComissaoGenerica.dValor, "Standard")
        objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_PercentualEmissao_Col) = Format(objComissaoGenerica.dPercentualEmissao, "Percent")
        objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorEmissao_Col) = Format(objComissaoGenerica.dValorEmissao, "Standard")
        objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_PercentualBaixa_Col) = Format(1 - objComissaoGenerica.dPercentualEmissao, "Percent")
        objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorBaixa_Col) = Format(objComissaoGenerica.dValor - objComissaoGenerica.dValorEmissao, "Standard")
        
        '*** 15/04/02 - INÍCIO Luiz G.F.Nogueira ***
        'Descobre o valor que será exibido no campo DiretoIndireto
        Select Case objComissaoGenerica.iIndireta
        
            'Se a comissão foi para um vendedor direto
            Case VENDEDOR_DIRETO
            
                'Guarda o texto "Direto"
                sDiretoIndireto = VENDEDOR_DIRETO_STRING
            
            'Se a comissão for para um vendedor indireto
            Case VENDEDOR_INDIRETO

                'Guarda o texto "Indireto"
                sDiretoIndireto = VENDEDOR_INDIRETO_STRING
            
        End Select
        
        'Exibe no grid o text encontrado para o campo DiretoIndireto
        objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_DiretoIndireto_Col) = sDiretoIndireto
        '*** 15/04/02 - FIM Luiz G.F.Nogueira ***
        
    Next

    'Atualiza o numero de linhas existentes apos o preenchimento de todas as linhas
    objTela.objGridComissoes.iLinhasExistentes = iIndice

    'Tulio-10/05/02
    Call Calcula_Valores
    
    Carrega_Grid_Comissoes = SUCESSO

    Exit Function

Erro_Carrega_Grid_Comissoes:

    Carrega_Grid_Comissoes = gErr

    Select Case gErr

        Case 98983

        Case 98984
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", gErr, objVendedor.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153586)

    End Select

End Function

Public Function Comissoes_Calcula_Padrao() As Long
'Exibe comissoes automaticas quando é recalculado valor total da NF/PedVenda
'ou quando clicka em Comissoes Automatica checkbox

'Nome alterado por tulio em 20/03/02, De Comissoes_Automatica_Exibe para Comissoes_Calcula_Padrao

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim objComissaoNF As New ClassComissaoNF
Dim iIndice As Integer

On Error GoTo Erro_Comissoes_Calcula_Padrao

    'Se existir linha no grid de comissões
    If objTela.Controls("ComissaoAutomatica").Value = vbChecked And objTela.objGridComissoes.iLinhasExistentes > 0 Then
        'para cada comissão no grid
        For iIndice = 1 To objTela.objGridComissoes.iLinhasExistentes

            'Se vendedor e percentuais estão preenchidos,
            If Len(Trim(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_Vendedor_Col))) > 0 And Len(Trim(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_PercentualComissao_Col))) > 0 And Len(Trim(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_PercentualEmissao_Col))) > 0 Then

                objVendedor.sNomeReduzido = objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_Vendedor_Col)

                'Lê o Vendedor
                lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
                If lErro <> SUCESSO And lErro <> 25008 Then Error 28563

                'Se não achou o nome do Vendedor --> erro
                If lErro <> SUCESSO Then Error 28656

                'Pega percentuais do Grid
                objComissaoNF.dPercentual = PercentParaDbl(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_PercentualComissao_Col))
                objComissaoNF.dPercentualEmissao = PercentParaDbl(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_PercentualEmissao_Col))

                'Obtem valores (base, comissao, comissao emissao)
                lErro = Comissao_Obter_Valores(objComissaoNF, objVendedor)
                If lErro <> SUCESSO Then Error 25726

                'Exibe na tela os valores da comissao
                objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorBase_Col) = Format(objComissaoNF.dValorBase, "Standard")
                objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorComissao_Col) = Format(objComissaoNF.dValor, "Standard")
                objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorEmissao_Col) = Format(objComissaoNF.dValorEmissao, "Standard")
                objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorBaixa_Col) = Format(objComissaoNF.dValor - objComissaoNF.dValorEmissao, "Standard")
                objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_DiretoIndireto_Col) = IIf(objComissaoNF.iIndireta <> VENDEDOR_DIRETO, VENDEDOR_INDIRETO_STRING, VENDEDOR_DIRETO_STRING)
            

            End If

        Next

        'Tulio-10/05/02
        Call Calcula_Valores
        
    End If

    Comissoes_Calcula_Padrao = SUCESSO

    Exit Function

Erro_Comissoes_Calcula_Padrao:

    Comissoes_Calcula_Padrao = Err

    Select Case Err

        Case 25726, 28563

        Case 28656
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", Err, objVendedor.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153587)

    End Select

    Exit Function

End Function

Public Function Inicializa_Grid_Comissoes(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Comissões

Dim iIndice As Integer
Dim objFatGlobal As New ClassFatGlobal
Dim iEditaComissao As Integer
    
    If Me.objTela.Name = "PedidoVenda" Then
        iEditaComissao = objFatGlobal.gobjFAT.iPedidoVendaEditaComissao
    Else
        iEditaComissao = objFatGlobal.gobjFAT.iNFiscalEditaComissao
    End If
    
    'Se o sitema estiver configurado para editar comissão
    If iEditaComissao = MARCADO Then
        'Habilita os controles que fazem parte da comissão
        objTela.Controls("Vendedor").Enabled = True
        objTela.Controls("PercentualComissao").Enabled = True
        objTela.Controls("ValorBase").Enabled = True
        objTela.Controls("ValorComissao").Enabled = True
        objTela.Controls("PercentualEmissao").Enabled = True
        objTela.Controls("ValorEmissao").Enabled = True
        objTela.Controls("ComissaoAutomatica").Enabled = True
        objTela.Controls("BotaoVendedores").Enabled = True
        objTela.Controls("DiretoIndireto").Enabled = True
    
    'Senão
    Else
        'Desabilita os controles que fazem parte da comissão do Pedido
        objTela.Controls("Vendedor").Enabled = False
        objTela.Controls("PercentualComissao").Enabled = False
        objTela.Controls("ValorBase").Enabled = False
        objTela.Controls("ValorComissao").Enabled = False
        objTela.Controls("PercentualEmissao").Enabled = False
        objTela.Controls("ValorEmissao").Enabled = False
        objTela.Controls("ComissaoAutomatica").Enabled = False
        objTela.Controls("BotaoVendedores").Enabled = False
        objTela.Controls("DiretoIndireto").Enabled = False
    End If

    Set objGridInt.objForm = objTela

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Vendedor")
    objGridInt.colColuna.Add ("Percentual")
    objGridInt.colColuna.Add ("Valor Base")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("% Emissão")
    objGridInt.colColuna.Add ("Valor Emissão")
    objGridInt.colColuna.Add ("% Baixa")
    objGridInt.colColuna.Add ("Valor Baixa")
    objGridInt.colColuna.Add ("Direta/Indireta")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (objTela.Controls("Vendedor").Name)
    objGridInt.colCampo.Add (objTela.Controls("PercentualComissao").Name)
    objGridInt.colCampo.Add (objTela.Controls("ValorBase").Name)
    objGridInt.colCampo.Add (objTela.Controls("ValorComissao").Name)
    objGridInt.colCampo.Add (objTela.Controls("PercentualEmissao").Name)
    objGridInt.colCampo.Add (objTela.Controls("ValorEmissao").Name)
    objGridInt.colCampo.Add (objTela.Controls("PercentualBaixa").Name)
    objGridInt.colCampo.Add (objTela.Controls("ValorBaixa").Name)
    objGridInt.colCampo.Add (objTela.Controls("DiretoIndireto").Name)

    iGrid_Vendedor_Col = 1
    iGrid_PercentualComissao_Col = 2
    iGrid_ValorBase_Col = 3
    iGrid_ValorComissao_Col = 4
    iGrid_PercentualEmissao_Col = 5
    iGrid_ValorEmissao_Col = 6
    iGrid_PercentualBaixa_Col = 7
    iGrid_ValorBaixa_Col = 8
    iGrid_DiretoIndireto_Col = 9

    'Grid do GridInterno
    objGridInt.objGrid = objTela.Controls("GridComissoes")

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_COMISSOES

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    objTela.Controls("GridComissoes").ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Comissoes = SUCESSO

    Exit Function

End Function

Public Function BotaoVendedores_Click() As Long

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection
Dim lErro As Long

On Error GoTo Erro_BotaoVendedores_Click
    
    If objTela.Controls("GridComissoes").Row = 0 Then Error 25733

    If Len(Trim(objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_Vendedor_Col))) > 0 Then

        objVendedor.sNomeReduzido = objTela.Controls("GridComissoes").TextMatrix(objTela.Controls("GridComissoes").Row, iGrid_Vendedor_Col)

    End If

    'Chama a tela de browse de Vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objTela.objEventoVendedor)

    BotaoVendedores_Click = SUCESSO
    
    Exit Function
    
Erro_BotaoVendedores_Click:

    BotaoVendedores_Click = Err
    
    Select Case Err
        
        Case 25733
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153588)
        
    End Select
    
    Exit Function

End Function

Public Function Move_TabComissoes_Memoria(objObjetoTela As Object, iTipoTela As Integer) As Long
'Recolhe os dados do Tab de Comissões
'objObjetoTela RETORNA(Output) as comissões dentro de uma coleção
'iTipoTela RECEBE(Input) o tipo da tela para indicar como objComissao e colComissoes devem ser instanciados

Dim lErro As Long
Dim objComissao As Object
Dim iIndice As Integer
Dim objVendedor As New ClassVendedor
Dim colComissoes As New Collection
Dim lTamanho As Long
Dim bCalculou As Boolean

On Error GoTo Erro_Move_TabComissoes_Memoria
    
    '*** 11/04/02 - INÍCIO Luiz G.F.Nogueira ***
    'Se a empresa usa as regras para calcular comissões
    If gobjCRFAT.iUsaComissoesRegras = USA_REGRAS Then
    
        '*** 15/04/02 - INÍCIO Luiz G.F.Nogueira (inclusão de bCalculou) ***
        'Calcula as comissões a partir das regras cadastradas
        lErro = objTela.objComissoesRegrasCalc.Comissoes_Calcula_Regras_Int(objTela, bCalculou, colComissoes)
        If lErro <> SUCESSO Then gError 102030
        '*** 15/04/02 - FIM Luiz G.F.Nogueira ***
    
        'Se calculou novas comissões, as comissões do grid devem ser descartadas,
        'e as comissões que estão na coleção são as que deverão ser gravadas no BD
        'Portanto, sai da função
        If bCalculou Then
            
            Set gColComissoes = colComissoes
        
            '*** 15/04/02 - INÍCIO Luiz G.F.Nogueira ***
            'Valida as comissões que foram calculadas
            lErro = Valida_Comissoes(colComissoes)
            
            'Se as comissões na coleção não são válidas
            If lErro <> SUCESSO Then
            
                'alteracao por tulio050203
                'para obrigar que, caso a validacao tenha dado erro,
                'da proxima vez q o sistema for mover as comissoes para a
                'memoria, ele tenha q recalcular as comissoes
                objTela.iComissoesAlterada = REGISTRO_ALTERADO
            
                'Exibe-as na tela, para que o usuário possa verificar o que está errado
                lErro = Carrega_Grid_Comissoes(colComissoes)
                If lErro <> SUCESSO Then gError 102043
                
                'Erro
                gError 102037
            
            End If
            '*** 15/04/02 - FIM Luiz G.F.Nogueira ***
            
        End If
    
    End If
    '*** 11/04/02 - FIM Luiz G.F.Nogueira ***
    
    'Se anida não calculou as comissões
    If Not bCalculou Then
    
        'para cada linha do grid
        If objTela.objGridComissoes.iLinhasExistentes > 0 Then
    
            'Recolhe as Comissões do Grid
            For iIndice = 1 To objTela.objGridComissoes.iLinhasExistentes
        
                'se a tela for de nota fiscal, instancia o obj como sendo da classe comissaonf
                If iTipoTela = NOTA_FISCAL Then
                    
                    Set objComissao = New ClassComissaoNF
                    
                'se a tela for de pedido de venda, instancia o dito cujo como sendo da classe
                'pedido de venda...
                ElseIf iTipoTela = PEDIDO_DE_VENDA Then
                
                    Set objComissao = New ClassComissaoPedVendas
                
                End If
                
                objComissao.iSeq = iIndice
        
                'se vendedor estiver preenchido
                If Len(Trim(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_Vendedor_Col))) > 0 Then
                    
                    'copia o nome reduzido do vendedor para o obj
                    objVendedor.sNomeReduzido = objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_Vendedor_Col)
        
                    'Lê o Vendedor, caso nao ache, pode haver inconsistencia
                    lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
                    If lErro <> SUCESSO And lErro <> 25008 Then Error 25735
                    If lErro = 25008 Then Error 25736
                    
                    'guarda no obj o codigo do vendedor..
                    objComissao.iCodVendedor = objVendedor.iCodigo
                
                End If
        
                'Armazena comissoes em memória
                lTamanho = Len(Trim(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_PercentualComissao_Col)))
                
                objComissao.dPercentual = IIf(lTamanho > 0, PercentParaDbl(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_PercentualComissao_Col)), 0)
        
                objComissao.dValorBase = StrParaDbl(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorBase_Col))
                
                objComissao.dValor = StrParaDbl(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorComissao_Col))
                
                objComissao.dValorEmissao = StrParaDbl(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorEmissao_Col))
        
                lTamanho = Len(Trim(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_PercentualEmissao_Col)))
                
                objComissao.dPercentualEmissao = IIf(lTamanho > 0, PercentParaDbl(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_PercentualEmissao_Col)), 0)
        
                objComissao.iIndireta = IIf(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_DiretoIndireto_Col) = VENDEDOR_DIRETO_STRING, VENDEDOR_DIRETO, VENDEDOR_INDIRETO)
        
                'Adiciona na coleção de comissões
                colComissoes.Add objComissao
                
                If iIndice <= gColComissoes.Count Then Set objComissao.colLinhas = gColComissoes.Item(iIndice).colLinhas
            
            Next
    
        End If
    
    End If

    'seleciona o tipo da tela passado
    Select Case iTipoTela
    
        'se for uma nota fiscal
        Case NOTA_FISCAL
            
            'instancia a colecao como a colecao de comissoesnf
            Set objObjetoTela.ColComissoesNF = colComissoes
                    
        'se for um pedido de venda
        Case PEDIDO_DE_VENDA
        
            'instancia a colecao como a colecao de comissoesnf
            Set objObjetoTela.colComissoes = colComissoes
        
            'move o valor da check para o obj
            objObjetoTela.iComissaoAutomatica = objTela.Controls("ComissaoAutomatica").Value
    
        'se for tela nao definida, sai por erro...
        Case Else
            gError 98985
    
    End Select
    
    Move_TabComissoes_Memoria = SUCESSO

    Exit Function

Erro_Move_TabComissoes_Memoria:

    Move_TabComissoes_Memoria = gErr

    Select Case gErr

        Case 25735, 102030, 102037  'tratado na rotina chamada
        
        Case 25736
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", gErr, objVendedor.sNomeReduzido)

        Case 98985
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TELA_INDEFINIDO", gErr)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153589)

    End Select

    Exit Function

End Function

Public Function Comissoes_Calcula() As Long

Dim lErro As Long

On Error GoTo Erro_Comissoes_Calcula

    'Se estiver carregando ou limpando a tela => sai da função pois não tem q haver cálculo de comissões no momento do carregamento
    If objTela.gbCarregandoTela Or objTela.gbLimpandoTela Then Exit Function

    If objTela.iComissoesAlterada <> 0 Then
    
        'se a empresa nao utiliza as regras p/ calculo de comissoes
        If Not gobjCRFAT.iUsaComissoesRegras = USA_REGRAS Then
        
            'Calcular as comissoes usando o modelo antigo
            lErro = Comissoes_Calcula_Padrao()
            If lErro <> SUCESSO Then gError 98978
        
        Else
        
            'Calcular as comissoes usando o novo modelo
            lErro = Comissoes_Calcula_Regras()
            If lErro <> SUCESSO Then gError 98979
        
        End If
    
        '*** 11/04/02 - Luiz G.F.Nogueira ***
        'Limpa iComissoesAlterada
        objTela.iComissoesAlterada = 0
        '************************************
    
    End If
    
    Comissoes_Calcula = SUCESSO
    
    Exit Function

Erro_Comissoes_Calcula:

    Comissoes_Calcula = gErr

    Select Case gErr
    
        Case 98978, 98979
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153590)
            
    End Select
    
End Function

Public Function Comissoes_Calcula_Regras() As Long
'Chama a rotina que ira calcular as comissoes utilizando  as regras cadastradas

Dim lErro As Long
Dim colComissoes As New Collection
Dim bCalculou As Boolean
 
On Error GoTo Erro_Comissoes_Calcula_Regras

    'Permite o redesenho da tela
    DoEvents
    
    'Calcula as comissoes com as regras
    lErro = objTela.objComissoesRegrasCalc.Comissoes_Calcula_Regras_Int(objTela, bCalculou, colComissoes)
    If lErro <> SUCESSO Then gError 98980

    '*** 15/04/02 - INÍCIO Luiz G.F.Nogueira ***
    'Se não foi executado o cálculo de comissões => sai da função, pois não é preciso carregar o grid
    If Not bCalculou Then Exit Function
    '*** 15/04/02 - FIM Luiz G.F.Nogueira ***
    
    'Exibe as comissoes no grid
    lErro = Carrega_Grid_Comissoes(colComissoes)
    If lErro <> SUCESSO Then gError 98981
    
    Set gColComissoes = colComissoes
    
    Comissoes_Calcula_Regras = SUCESSO

    Exit Function

Erro_Comissoes_Calcula_Regras:

    Comissoes_Calcula_Regras = gErr

    Select Case gErr
        
        Case 98980, 98981
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153591)
        
    End Select

    Exit Function

End Function

'*** 15/04/02 - INÍCIO Luiz G.F.Nogueira ***
Public Function Valida_Comissoes(ByVal colComissoes As Collection) As Long
'Verifica se as comissões que estã na coleção são válidas para gravação
'colComissoes RECEBE(Input) as comissões que devem ser validadas

Dim objComissao As Object
Dim dValorTotalComissao As Double
Dim dValorTotal As Double
Dim iIndice As Integer

On Error GoTo Erro_Valida_Comissoes

    'Para cada comissão na coleção
    For Each objComissao In colComissoes
    
        'Incrementa iIndice (essa variável serve apenas para permitir que a mensagem de erro indique a linha
        'onde existe um campo com problemas)
        iIndice = iIndice + 1
        
        'Se não foi preenchido o vendedor => erro
        If objComissao.iCodVendedor = 0 Then gError 102038
        
        'Se o valor da comissão não foi preenchido => erro
        If objComissao.dValor = 0 Then gError 102039
        
        '---- trecho comentado por tulio050203
        'Se o valor da comissão a ser pago na emissão não foi preenchido => erro
        'If objComissao.dValorEmissao = 0 Then gError 102040
        
        'Se o percentual da comissão a ser pago na emissão não foi preenchido => erro
        'If objComissao.dPercentualEmissao = 0 Then gError 102041
        '---- fim trecho comentado por tulio050203
        
        'Acumula o valor da comissão
        dValorTotalComissao = dValorTotalComissao + objComissao.dValor
    
    Next

    'Guarda o valor total da nota que está na tela
    dValorTotal = StrParaDbl(objTela.Controls("ValorTotal").Caption)

    'Se o valor total de comissões é maior do que o valor total da tela
    If dValorTotalComissao > dValorTotal Then gError 102042
    
    Valida_Comissoes = SUCESSO

    Exit Function

Erro_Valida_Comissoes:

    Valida_Comissoes = gErr

    Select Case gErr

        Case 102040
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_EMISSAO_GRID_NAO_PREENCHIDO", gErr, iIndice)
            
        Case 102041
            Call Rotina_Erro(vbOKOnly, "ERRO_PERCENTAGEM_EMISSAO_GRID_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 102038
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_COMISSAO_NAO_INFORMADO", gErr, iIndice)

        Case 102039
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_COMISSAO_GRID_NAO_PREENCHIDO", gErr, iIndice)

        Case 102042
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TOTAL_COMISSAO_INVALIDO", gErr, dValorTotalComissao, dValorTotal)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153592)

    End Select

End Function
'*** 15/04/02 - FIM Luiz G.F.Nogueira ***

Private Function Calcula_Valores() As Long
'Tulio-10/05/02
'Essa funcao tem por objetivo calcular os valores (valorbase, percentual, valor)
'e coloca-los nas labels das telas
'Existem 2 funcoes comentadas no fim desse codigo, sao as antigas que nao estao muito bem feitas
'mas servem de backup...

Dim iIndice As Integer
Dim dSomaPercentual As Double
Dim dValor As Double
Dim dValorBase As Double

On Error GoTo Erro_Calcula_Valores

    'Loop no GridComissao
    For iIndice = 1 To objTela.objGridComissoes.iLinhasExistentes

        'Acumula Valor da comissao em dValor
        dValor = dValor + StrParaDbl(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorComissao_Col))

        'Acumula o ValorBase em dValorBase
        'dValorBase = dValorBase + StrParaDbl(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorBase_Col))

    Next

    'alteracao por tulio 27/06
    dValorBase = StrParaDbl(objTela.Controls("ValorTotal").Caption)
    
    'se o ValorBase for diferente de 0
    If dValorBase <> 0 Then

        'o percentual eh calculado a partir da razao entre o somatorio de dvalor e o somatorio
        'de valorbase
        dSomaPercentual = dValor / dValorBase

    End If

    'Mostra na tela o Total Percentual do somatorio de valor sobre o somatorio de valorbase
    objTela.Controls("TotalPercentualComissao").Caption = Format(dSomaPercentual, "Percent")
    
    'Mostra na tela o Total Valor
    objTela.Controls("TotalValorComissao").Caption = Format(dValor, "Standard")

    'Mostra na tela o Total do ValorBase
    objTela.Controls("TotalValorBase").Caption = Format(dValorBase, "Standard")

    Calcula_Valores = SUCESSO

    Exit Function

Erro_Calcula_Valores:

    Calcula_Valores = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153593)

    End Select

    Exit Function

End Function


'Comentadas por Tulio-10/05/02
'Remover apos testes com tab comissoes...
Public Function Soma_Percentual() As Long

Dim iIndice As Integer
Dim dSomaPercentual As Double
Dim dValor As Double
Dim dValorBase As Double
Dim lComprimento As Long
Dim lErro As Long

On Error GoTo Erro_Soma_Percentual

    '??? retirar
    'inicializa as variaveis
    dSomaPercentual = 0
    dValorBase = 0
    dValor = 0

    'Loop no GridComissao
    For iIndice = 1 To objTela.objGridComissoes.iLinhasExistentes

        '??? comentar
        lComprimento = Len(Trim(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorComissao_Col)))

        'Verifica se Valor da Comissão está preenchido
        If lComprimento > 0 Then

            'Acumula Valor da comissao em dValor
            dValor = dValor + StrParaDbl(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorComissao_Col))

        End If

        '??? comentar
        lComprimento = Len(Trim(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorBase_Col)))

        'Verifica se Valor Base da Comissão está preenchido
        If lComprimento > 0 Then

            'Acumula ValorBase em dValorBase
            dValorBase = dValorBase + StrParaDbl(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorBase_Col))

        End If

    Next

    'se valor base for 0, soma percentual sempre eh 0%
    If dValorBase = 0 Then

        dSomaPercentual = 0

    Else
    'se valorbase nao for 0, o percentual varia num intervalo fechado entre 0% e 100%

        dSomaPercentual = dValor / dValorBase

    End If

    'Mostra na tela o Total Percentual do somatorio de valor sobre o somatorio de valorbase
    objTela.Controls("TotalPercentualComissao").Caption = Format(dSomaPercentual, "Percent")

    Soma_Percentual = SUCESSO

    Exit Function

Erro_Soma_Percentual:

    Soma_Percentual = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153594)

    End Select

    Exit Function

End Function

'Funcoes inuteis, remover apos testes caso nao de problema
'tulio-10/05/02
Public Function Soma_Valor() As Long

Dim iIndice As Integer
Dim dSomaValor As Double
Dim lErro As Long

On Error GoTo Erro_Soma_Valor

    dSomaValor = 0

    'Loop no GridComissao
    For iIndice = 1 To objTela.objGridComissoes.iLinhasExistentes

        'Verifica se Valor da Comissão está preenchido
        If Len(Trim(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorComissao_Col))) > 0 Then

            'Acumula Valor em dSomaValor
            dSomaValor = dSomaValor + CDbl(objTela.Controls("GridComissoes").TextMatrix(iIndice, iGrid_ValorComissao_Col))

        End If

    Next

    'Mostra na tela o Total Valor
    objTela.Controls("TotalValorComissao").Caption = Format(dSomaValor, "Standard")

    Soma_Valor = SUCESSO

    Exit Function

Erro_Soma_Valor:

    Soma_Valor = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153595)

    End Select

    Exit Function

End Function

Private Function Obtem_Tributacao_Tela(objTribDoc As Object) As Long

Dim objTribItem As Object
Dim colTribItens As Collection

On Error GoTo Erro_Obtem_Tributacao_Tela
    
    If Existe_objTribTab_Tela Then
        
        Set objTribDoc = objTela.gobjTribTab.gobjDocTribTributacao
        Set colTribItens = objTela.gobjTribTab.gcolDocTribItem
    
        objTribDoc.dPISValor = 0
        objTribDoc.dCOFINSValor = 0
        
        For Each objTribItem In colTribItens
            objTribDoc.dPISValor = objTribDoc.dPISValor + objTela.gobjTribTab.gobjDocTribItemTributacao(objTribItem).dPISValor
            objTribDoc.dCOFINSValor = objTribDoc.dCOFINSValor + objTela.gobjTribTab.gobjDocTribItemTributacao(objTribItem).dCOFINSValor
        Next
    
        objTribDoc.dPISValor = Arredonda_Moeda(objTribDoc.dPISValor)
        objTribDoc.dCOFINSValor = Arredonda_Moeda(objTribDoc.dCOFINSValor)
    
    Else
        Set objTribDoc = New ClassTributacaoDoc
    End If

    Obtem_Tributacao_Tela = SUCESSO

    Exit Function

Erro_Obtem_Tributacao_Tela:

    Obtem_Tributacao_Tela = Err

    Exit Function
    
End Function

Private Function Obtem_ValorItens_NaoFat(dValor As Double) As Long

Dim lErro As Long
Dim objTribItem As Object
Dim bFat As Boolean, iItem As Integer, dValorAux As Double

On Error GoTo Erro_Obtem_ValorItens_NaoFat

    dValor = 0
    
    If Existe_objTribTab_Tela Then

        For Each objTribItem In objTela.gobjTribTab.gcolDocTribItem
            
            iItem = iItem + 1

            If objTela.gobjTribTab.iTipoTela = objTela.gobjTribTab.TIPOTELA_NF Then
    
                lErro = objTela.gobjTribTab.Testa_Item_Fat(iItem, bFat)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                If Not bFat Then
                    dValorAux = dValorAux + objTribItem.dPrecoTotal
                End If
    
            End If
            
        Next
    
    End If
    
    dValor = dValorAux

    Obtem_ValorItens_NaoFat = SUCESSO

    Exit Function

Erro_Obtem_ValorItens_NaoFat:

    Obtem_ValorItens_NaoFat = gErr
    
    dValor = 0

    Exit Function
    
End Function

Private Function Existe_objTribTab_Tela() As Boolean

On Error GoTo Erro_Existe_objTribTab_Tela

    Existe_objTribTab_Tela = False
    If Not (objTela.gobjTribTab.gobjDocTribTributacao Is Nothing) Then
        Existe_objTribTab_Tela = True
    End If

    Exit Function

Erro_Existe_objTribTab_Tela:

    Existe_objTribTab_Tela = False

    Exit Function
    
End Function

Public Function VendedorLinha_Preenche2(objVendedor As ClassVendedor) As Long
'preenche a linha corrente do grid de comissoes p/o vendedor passado com valores default

Dim lErro As Long, iIndice As Integer
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_VendedorLinha_Preenche2

    'chama rotina que ira verificar se o vendedor ja esta no grid
    lErro = VendedorLinha_Preenche1(objVendedor)
    If lErro <> SUCESSO Then gError 98989
    
    'Exibe comissão automática vinculada ao Vendedor
    lErro = Comissao_Automatica_Vendedor_Exibe(objTela.Controls("GridComissoes").Row, objFilialCliente, objVendedor)
    If lErro <> SUCESSO Then Error 25693

    'Tulio-10/05/02
    Call Calcula_Valores

    VendedorLinha_Preenche2 = SUCESSO

    Exit Function

Erro_VendedorLinha_Preenche2:

    VendedorLinha_Preenche2 = gErr

    Select Case gErr

        Case 25692, 25693, 98989 'tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153574)

    End Select

    Exit Function

End Function

Public Sub ExcluiLinhaGrid(ByVal iLinha As Integer)

Dim objComissao As Object, iIndice As Integer

    Call Soma_Percentual
    Call Soma_Valor

    If gColComissoes.Count >= iLinha Then
    
        gColComissoes.Remove (iLinha)
        
        For Each objComissao In gColComissoes
        
            iIndice = iIndice + 1
            objComissao.iSeq = iIndice
        
        Next
        
    End If
    
End Sub

Public Sub Limpa_Grid_Comissoes()

    Call Grid_Limpa(objTela.objGridComissoes)
    
    Set gColComissoes = New Collection
        
End Sub
