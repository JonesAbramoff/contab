VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCRFATCritica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function PorcentagemJuros_Critica(ByVal sString As String) As Long
'Critica o valor passado em sString , se nao for entre 0 e PORCENTAGEM_JUROS_MAXIMA retorna erro.

Dim dValor As Double
Dim lErro As Long

On Error GoTo Error_PorcentagemJuros_Critica

    If IsNumeric(sString) Then
        dValor = 0#
        dValor = CDbl(sString)
    Else
        Error 25088
    End If

    If dValor > PORCENTAGEM_JUROS_MAXIMA Or dValor < 0 Then Error 25089

    PorcentagemJuros_Critica = SUCESSO

    Exit Function

Error_PorcentagemJuros_Critica:

    PorcentagemJuros_Critica = Err

    Select Case Err

        Case 25088
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", Err, sString)
            
        Case 25089
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_PORCENTAGEM_JUROS", Err, dValor, PORCENTAGEM_JUROS_MAXIMA)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148816)

    End Select

    Exit Function

End Function

Function Parcelas_Calcula(ByVal dValor As Double, ByVal iNumParcelas As Integer, ByVal colValorParcelas As Collection) As Long
'Calcula os valores das Parcelas a partir do Valor e Número de Parcelas
'Devolve em colValorParcelas

Dim iIndice As Integer
Dim iResto As Integer
Dim dValorParcela As Double

    If iNumParcelas <> 1 Then
    
        'Calcula o resto da divisão inteira
        iResto = Resto(dValor * 100, iNumParcelas)
    
        If (iResto <> 0) Then
            dValorParcela = (dValor * 100) / iNumParcelas
            dValorParcela = Int(dValorParcela) / 100
        Else
            dValorParcela = dValor / iNumParcelas
        End If
        
    Else
    
        dValorParcela = dValor
        iResto = 0
        
    End If

    'Acrescentar valores das parcelas na coleção
    For iIndice = 1 To iNumParcelas - iResto

        colValorParcelas.Add dValorParcela

    Next

    'Soma 0.01 ao Valor da Parcela
    dValorParcela = dValorParcela + 0.01

    'Se a divisão não foi exata acrescentar as "iResto" últimas parcelas adicionadas de 0.01
    For iIndice = 1 To iResto

        colValorParcelas.Add dValorParcela

    Next

    Parcelas_Calcula = SUCESSO

    Exit Function

End Function

Function Parcelas_DatasVencimento(ByVal objCondicaoPagto As ClassCondicaoPagto, ByVal dtDataEmissao As Date, ByVal colDataVencimento As Collection) As Long
'Calcula as datas de vencimento de cada uma das Parcelas do Grid

Dim lErro As Long
Dim dtDataVencimento As Date
Dim iIndice As Integer
Dim iDiferenca As Integer

    'Se a condição for mensal
    If objCondicaoPagto.iMensal <> False And objCondicaoPagto.iNumeroParcelas > 1 Then

        'Calcula a data com intervalo para a primeira parcela
        dtDataVencimento = dtDataEmissao + objCondicaoPagto.iDiasParaPrimeiraParcela

        'Se o dia dessa data for menor que o dia do mês de pagamento da Parcela incrementa o mês
        iIndice = 1

        Do While iIndice <= objCondicaoPagto.iNumeroParcelas

            If objCondicaoPagto.iDiaDoMes > Day(dtDataVencimento) Then
                Do While (Month(dtDataVencimento + 1) = Month(dtDataVencimento)) And (Day(dtDataVencimento) <> objCondicaoPagto.iDiaDoMes)
                    dtDataVencimento = dtDataVencimento + 1
                Loop
            Else
                iDiferenca = Abs(Day(dtDataVencimento - objCondicaoPagto.iDiaDoMes))
                dtDataVencimento = dtDataVencimento - iDiferenca
                dtDataVencimento = DateAdd("m", 1, dtDataVencimento)
            
            End If
            
            'Guarda a data de vencimento em colDataVencimento
            colDataVencimento.Add dtDataVencimento

            dtDataVencimento = DateAdd("m", 1, dtDataVencimento)
            iIndice = iIndice + 1
        Loop

    Else 'Se a condição não for mensal

        'Calcula a primeira data de vencimento
        dtDataVencimento = dtDataEmissao + objCondicaoPagto.iDiasParaPrimeiraParcela

        'Loop para acrescentar as datas na coleção
        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas
            colDataVencimento.Add dtDataVencimento
            dtDataVencimento = dtDataVencimento + objCondicaoPagto.iIntervaloParcelas
        Next

    End If

    Parcelas_DatasVencimento = SUCESSO

    Exit Function

End Function

Function Parcela_GeraDescontoPadrao(ByVal colDesconto As colDesconto, ByVal dtVencimentoParc As Date) As Long
'preenche a colecao de descontos padrao para a parcela

Dim lErro As Long, colDescontoPadrao As New Collection
Dim objDescontoPadrao As ClassDescontoPadrao

On Error GoTo Erro_Parcela_GeraDescontoPadrao

    If dtVencimentoParc <> DATA_NULA Then
    
        lErro = CF("DescontoPadrao_Le", colDescontoPadrao)
        If lErro <> SUCESSO Then Error 56584
    
        For Each objDescontoPadrao In colDescontoPadrao
        
            'Se os atributos de objDescontoPadrao padrão estiverem preenchidos adiciona em coldesconto
            If objDescontoPadrao.iCodigo > 0 And objDescontoPadrao.dPercentual > 0 Then
                
                colDesconto.Add objDescontoPadrao.iCodigo, dtVencimentoParc - objDescontoPadrao.iDias, objDescontoPadrao.dPercentual
                
            End If
        
        Next

    End If
    
    Parcela_GeraDescontoPadrao = SUCESSO
     
    Exit Function
    
Erro_Parcela_GeraDescontoPadrao:

    Parcela_GeraDescontoPadrao = Err
     
    Select Case Err
          
        Case 56584
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148817)
     
    End Select
     
    Exit Function

End Function

Function Calcula_Desconto_Parcela(ByVal objParcelaReceber As ClassParcelaReceber, dValorDesconto As Double, ByVal dtDataBaixa As Date) As Long
'Obtem o valor do desconto por antecipacao de pagamento de uma parcela na data informada

Dim lErro As Long
Dim objDesconto As New ClassDesconto
Dim iDiasUteis As Integer
Dim iDias As Integer

On Error GoTo Erro_Calcula_Desconto_Parcela

    iDias = objParcelaReceber.dtDataVencimento - dtDataBaixa

    objDesconto.dtData = DATA_NULA
    dValorDesconto = 0
            
    If objParcelaReceber.dSaldo = objParcelaReceber.dValor And iDias >= 0 Then
        
        If objParcelaReceber.dtDesconto1Ate <> DATA_NULA Then
            
            If objParcelaReceber.dtDesconto1Ate >= dtDataBaixa Then
                objDesconto.dtData = objParcelaReceber.dtDesconto1Ate
                objDesconto.dValor = objParcelaReceber.dDesconto1Valor
                objDesconto.iCodigo = objParcelaReceber.iDesconto1Codigo
            End If
        End If
        
        If objParcelaReceber.dtDesconto2Ate <> DATA_NULA Then
            If objParcelaReceber.dtDesconto2Ate >= dtDataBaixa Then
                If objDesconto.dtData <> DATA_NULA Then
                    If objDesconto.dtData > objParcelaReceber.dtDesconto2Ate Then
                        objDesconto.dtData = objParcelaReceber.dtDesconto2Ate
                        objDesconto.dValor = objParcelaReceber.dDesconto2Valor
                        objDesconto.iCodigo = objParcelaReceber.iDesconto2Codigo
                    End If
                Else
                    objDesconto.dtData = objParcelaReceber.dtDesconto2Ate
                    objDesconto.dValor = objParcelaReceber.dDesconto2Valor
                    objDesconto.iCodigo = objParcelaReceber.iDesconto2Codigo
                End If
            End If
        End If
        
        If objParcelaReceber.dtDesconto3Ate <> DATA_NULA Then
            If objParcelaReceber.dtDesconto3Ate >= dtDataBaixa Then
                If objDesconto.dtData <> DATA_NULA Then
                    If objDesconto.dtData > objParcelaReceber.dtDesconto3Ate Then
                        objDesconto.dtData = objParcelaReceber.dtDesconto3Ate
                        objDesconto.dValor = objParcelaReceber.dDesconto3Valor
                        objDesconto.iCodigo = objParcelaReceber.iDesconto3Codigo
                    End If
                Else
                    objDesconto.dtData = objParcelaReceber.dtDesconto3Ate
                    objDesconto.dValor = objParcelaReceber.dDesconto3Valor
                    objDesconto.iCodigo = objParcelaReceber.iDesconto3Codigo
                End If
            End If
        End If
    End If
    
    If objDesconto.dtData <> DATA_NULA Then
    
        Select Case objDesconto.iCodigo
        
            Case VALOR_FIXO
                dValorDesconto = objDesconto.dValor
                
            Case VALOR_ANT_DIA
                dValorDesconto = objDesconto.dValor * iDias
                
            Case VALOR_ANT_DIA_UTIL
            
                lErro = CF("Dias_Uteis", dtDataBaixa, objParcelaReceber.dtDataVencimento, iDiasUteis)
                If lErro <> SUCESSO Then Error 46509
                
                dValorDesconto = objDesconto.dValor * iDiasUteis
        
            Case Percentual
                dValorDesconto = objDesconto.dValor * objParcelaReceber.dValor
    
            Case PERC_ANT_DIA
                dValorDesconto = objDesconto.dValor * iDias * objParcelaReceber.dValor
            
            Case PERC_ANT_DIA_UTIL
                lErro = CF("Dias_Uteis", dtDataBaixa, objParcelaReceber.dtDataVencimento, iDiasUteis)
                If lErro <> SUCESSO Then Error 46510
                
                dValorDesconto = objDesconto.dValor * iDiasUteis * objParcelaReceber.dValor
            
        End Select
        
        dValorDesconto = CDbl(Format(dValorDesconto, "0.00"))
        If dValorDesconto > objParcelaReceber.dValor Then dValorDesconto = objParcelaReceber.dValor
    
    End If
    
    Calcula_Desconto_Parcela = SUCESSO
    
    Exit Function
    
Erro_Calcula_Desconto_Parcela:

    Calcula_Desconto_Parcela = Err
    
    Select Case Err
    
        Case 46509, 46510
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148818)
            
    End Select

    Exit Function

End Function

Function Calcula_Multa_Juros_Parcela(ByVal objParcelaReceber As ClassParcelaReceber, ByVal dtDataBaixa As Date, dValorMulta As Double, dValorJuros As Double) As Long
'Obtem o valor da multa e dos juros para pagto da parcela na data informada

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber
Dim iDias As Integer, dValorJurosDia As Double, dtDataVencimentoReal As Date

On Error GoTo Erro_Calcula_Multa_Juros_Parcela
    
    lErro = CF("DataVencto_Real", objParcelaReceber.dtDataVencimento, dtDataVencimentoReal)
    If lErro <> SUCESSO Then Error 56719
    
    If dtDataBaixa > dtDataVencimentoReal Then
                
        objTituloReceber.lNumIntDoc = objParcelaReceber.lNumIntTitulo
        
        lErro = CF("TituloReceber_Le", objTituloReceber)
        If lErro <> SUCESSO And lErro <> 26061 Then Error 56718
        If lErro <> SUCESSO Then Error 56719
                
        iDias = dtDataBaixa - objParcelaReceber.dtDataVencimento
        
        dValorMulta = objParcelaReceber.dSaldo * objTituloReceber.dPercMulta
        
        'com juros compostos
        'dValorJuros = objParcelaReceber.dSaldo * (((1 + objTituloReceber.dPercJurosDiario) ^ iDias) - 1)
        
        'com juros simples
        dValorJurosDia = Round(objParcelaReceber.dSaldo * objTituloReceber.dPercJurosDiario, 2)
        dValorJuros = Round(dValorJurosDia * iDias, 2)
        
    Else
               
        dValorMulta = 0
        dValorJuros = 0
                       
    End If

    Calcula_Multa_Juros_Parcela = SUCESSO
    
    Exit Function

Erro_Calcula_Multa_Juros_Parcela:

    Calcula_Multa_Juros_Parcela = Err
    
    Select Case Err
    
        Case 56718
        
        Case 56719
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO", Err, objTituloReceber.lNumIntDoc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148819)
    
    End Select
    
    Exit Function

End Function

Public Function ComissaoAutomatica_Obter_Parcela(ByVal objVendedor As ClassVendedor, ByVal dValorParcela As Double, ByVal objComissaoNF As ClassComissaoNF, ByVal objComissao As ClassComissao) As Long
'retorna em objComissao valores default para geracao de comissao automatica à partir do cliente, filial e valor da parcela informados
'se objVendedor.iCodigo <> 0 entao obter o vendedor padrao da filial cliente ou do tipo

Dim lErro As Long

On Error GoTo Erro_ComissaoAutomatica_Obter_Parcela

    objComissao.dValorBase = dValorParcela
    objComissao.dValor = 0
    objComissao.dPercentual = 0
    objComissao.iCodVendedor = 0
    
    'Verifica se achou o Vendedor
    If objVendedor.iCodigo <> 0 Then
    
        objComissao.iCodVendedor = objVendedor.iCodigo
        objComissao.dValor = CDbl(Format(objComissao.dValorBase * objComissaoNF.dPercentual * objVendedor.dPercComissaoBaixa, "0.00"))
        objComissao.dPercentual = CDbl(Format(objComissaoNF.dPercentual * objVendedor.dPercComissaoBaixa, "0.0000"))
        
    End If
    
    ComissaoAutomatica_Obter_Parcela = SUCESSO
     
    Exit Function
    
Erro_ComissaoAutomatica_Obter_Parcela:

    ComissaoAutomatica_Obter_Parcela = Err
     
    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148820)
     
    End Select
     
    Exit Function

End Function

Public Function ComissaoAutomatica_Obter_Emissao(ByVal sClienteNomeRed As String, ByVal iCodFilial As Integer, ByVal objVendedor As ClassVendedor, ByVal objComissao As ClassComissao, ByVal dValorTotal As Double, ByVal dValorFrete As Double, ByVal dValorSeguro As Double, ByVal dValorOutras As Double, ByVal dValorDesconto As Double) As Long
'retorna em objVendedor e objComissao valores default para geracao de comissao automatica à partir do cliente e filial informados
'se objVendedor.iCodigo = 0 entao obter o vendedor padrao da filial cliente ou do tipo

Dim lErro As Long, objComissaoNF As New ClassComissaoNF

On Error GoTo Erro_ComissaoAutomatica_Obter_Emissao

    objComissao.dValorBase = 0
    objComissao.dValor = 0
    objComissao.dPercentual = 0
    objComissao.iCodVendedor = 0
    
    lErro = CF("ComissaoAutomatica_Obter_Info", sClienteNomeRed, iCodFilial, objVendedor, objComissaoNF)
    If lErro <> SUCESSO Then Error 59004
    
    'Verifica se achou o Vendedor
    If objVendedor.iCodigo <> 0 Then
    
        objComissao.iCodVendedor = objVendedor.iCodigo
    
        If objVendedor.iComissaoSobreTotal = 0 Then
            objComissao.dValorBase = dValorTotal - IIf(objVendedor.iComissaoFrete = 0, dValorFrete, 0) - IIf(objVendedor.iComissaoSeguro = 0, dValorSeguro, 0) - IIf(objVendedor.iComissaoICM = 0, dValorOutras, 0)
        Else
            objComissao.dValorBase = dValorTotal
        End If
        
        objComissao.dValor = CDbl(Format(objComissao.dValorBase * objComissaoNF.dPercentual * objComissaoNF.dPercentualEmissao, "0.00"))
        objComissao.dPercentual = CDbl(Format(objComissaoNF.dPercentual * objVendedor.dPercComissaoEmissao, "0.0000"))
        
    End If
    
    ComissaoAutomatica_Obter_Emissao = SUCESSO
     
    Exit Function
    
Erro_ComissaoAutomatica_Obter_Emissao:

    ComissaoAutomatica_Obter_Emissao = Err
     
    Select Case Err
          
        Case 59004
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148821)
     
    End Select
     
    Exit Function

End Function

Public Function ComissaoAutomatica_Obter_Info(ByVal sClienteNomeRed As String, ByVal iCodFilial As Integer, ByVal objVendedor As ClassVendedor, ByVal objComissaoNF As ClassComissaoNF) As Long
'retorna em objVendedor e objComissao valores default para geracao de comissao automatica à partir do cliente e filial informados
'se objVendedor.iCodigo = 0 entao obter o vendedor padrao da filial cliente ou do tipo

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_ComissaoAutomatica_Obter_Info

    If iCodFilial <> 0 And sClienteNomeRed <> "" Then
    
        objFilialCliente.iCodFilial = iCodFilial
        
        'Lê FilialCliente
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sClienteNomeRed, objFilialCliente)
        If lErro <> SUCESSO Then Error 59005
                   
        lErro = CF("ComissaoAutomatica_Obter_Info2", objFilialCliente, objVendedor, objComissaoNF)
        If lErro <> SUCESSO Then Error 59012
    
    End If
    
    ComissaoAutomatica_Obter_Info = SUCESSO
     
    Exit Function
    
Erro_ComissaoAutomatica_Obter_Info:

    ComissaoAutomatica_Obter_Info = Err
     
    Select Case Err
          
        Case 59005, 59012

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148822)
     
    End Select
     
    Exit Function

End Function

Public Function ComissaoAutomatica_Obter_Info2(ByVal objFilialCliente As ClassFilialCliente, ByVal objVendedor As ClassVendedor, ByVal objComissaoNF As ClassComissaoNF) As Long
'retorna em objVendedor e objComissao valores default para geracao de comissao automatica à partir do cliente e filial informados
'se objVendedor.iCodigo = 0 entao obter o vendedor padrao da filial cliente ou do tipo

Dim lErro As Long
Dim objTipoCliente As New ClassTipoCliente
Dim objCliente As New ClassCliente

On Error GoTo Erro_ComissaoAutomatica_Obter_Info2

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
                            
            If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
                objCliente.lCodigoLoja = objFilialCliente.lCodCliente
            Else
                objCliente.lCodigo = objFilialCliente.lCodCliente
            End If
            
            'Lê o Cliente
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO And lErro <> 12293 Then Error 59006
            
            'Não achou o Cliente --> erro
            If lErro = 12293 Then Error 59007
                    
            If objCliente.iTipo <> 0 Then
                
                objTipoCliente.iCodigo = objCliente.iTipo
                
                'Lê o Tipo de Cliente
                lErro = CF("TipoCliente_Le", objTipoCliente)
                If lErro <> SUCESSO And lErro <> 19062 Then Error 59008
                
                'Não achou o Tipo de Cliente --> erro
                If lErro = 19062 Then Error 59009
                
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
        If lErro <> SUCESSO And lErro <> 12582 Then Error 59010
        
        'Não achou o Vendedor --> erro
        If lErro = 12582 Then Error 59011
            
    End If
    
    'Verifica se achou o Vendedor
    If objVendedor.iCodigo <> 0 Then
    
        'Se ainda nao achou o % de comissao usar o do Vendedor
        If objComissaoNF.dPercentual = 0 Then objComissaoNF.dPercentual = objVendedor.dPercComissao
    
        objComissaoNF.dPercentualEmissao = objVendedor.dPercComissaoEmissao
    
    End If
    
    ComissaoAutomatica_Obter_Info2 = SUCESSO
     
    Exit Function
    
Erro_ComissaoAutomatica_Obter_Info2:

    ComissaoAutomatica_Obter_Info2 = Err
     
    Select Case Err
          
        Case 59006, 59008, 59010
        
        Case 59007
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objFilialCliente.lCodCliente)
            
        Case 59009
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_CLIENTE_NAO_CADASTRADO", Err, objTipoCliente.iCodigo)
            
        Case 59011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", Err, objVendedor.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148823)
     
    End Select
     
    Exit Function

End Function

Public Function ComissaoAutomatica_Obter_NF(ByVal objFilialCliente As ClassFilialCliente, ByVal objVendedor As ClassVendedor, ByVal objComissaoNF As ClassComissaoNF, ByVal dValorTotal As Double, ByVal dValorFrete As Double, ByVal dValorSeguro As Double, ByVal dValorOutras As Double, ByVal dValorDesconto As Double) As Long
'retorna em objComissao valores default para geracao de comissao automatica
'se objVendedor.iCodigo = 0 entao obter o vendedor padrao da filial cliente ou do tipo

Dim lErro As Long

On Error GoTo Erro_ComissaoAutomatica_Obter_NF

    objComissaoNF.dValorBase = 0
    objComissaoNF.dValor = 0
    objComissaoNF.dValorEmissao = 0

    lErro = CF("ComissaoAutomatica_Obter_Info2", objFilialCliente, objVendedor, objComissaoNF)
    If lErro <> SUCESSO Then Error 59013

    'Verifica se achou o Vendedor
    If objVendedor.iCodigo <> 0 Then

        If objVendedor.iComissaoSobreTotal = 0 Then
            objComissaoNF.dValorBase = dValorTotal - IIf(objVendedor.iComissaoFrete = 0, dValorFrete, 0) - IIf(objVendedor.iComissaoSeguro = 0, dValorSeguro, 0) - IIf(objVendedor.iComissaoICM = 0, dValorOutras, 0)
        Else
            objComissaoNF.dValorBase = dValorTotal
        End If

        objComissaoNF.dValor = CDbl(Format(objComissaoNF.dValorBase * objComissaoNF.dPercentual, "0.00"))
        objComissaoNF.dValorEmissao = CDbl(Format(objComissaoNF.dValorBase * objComissaoNF.dPercentual * objComissaoNF.dPercentualEmissao, "0.00"))

    End If

    ComissaoAutomatica_Obter_NF = SUCESSO

    Exit Function

Erro_ComissaoAutomatica_Obter_NF:

    ComissaoAutomatica_Obter_NF = Err

    Select Case Err

        Case 59013

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148824)

    End Select

    Exit Function

End Function

Public Function Comissao_GridFAT_Valida(ByVal objGridComissoes As AdmGrid, ByVal iGrid_Vendedor_Col As Integer, ByVal iGrid_ValorCom_Col As Integer, ByVal dValorTotal As Double, ByVal dValorIRRF As Double) As Long
'verifica se o grid de comissoes está preenchido com valores válidos

Dim lErro As Long
Dim iIndice As Integer
Dim dTotalValorComissao As Double, dValorComissao As Double

On Error GoTo Erro_Comissao_GridFAT_Valida

    If objGridComissoes.iLinhasExistentes > 0 Then

        For iIndice = 1 To objGridComissoes.iLinhasExistentes

            If Len(Trim(objGridComissoes.objGrid.TextMatrix(iIndice, iGrid_Vendedor_Col))) = 0 Then Error 35929

            dValorComissao = StrParaDbl(objGridComissoes.objGrid.TextMatrix(iIndice, iGrid_ValorCom_Col))

            If dValorComissao = 0 Then Error 35930

            dTotalValorComissao = dTotalValorComissao + dValorComissao
            
        Next

        If dTotalValorComissao > (dValorTotal - dValorIRRF) Then Error 35931

    End If

    Comissao_GridFAT_Valida = SUCESSO
     
    Exit Function
    
Erro_Comissao_GridFAT_Valida:

    Comissao_GridFAT_Valida = Err
     
    Select Case Err
          
        Case 35929
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_COMISSAO_GRID_NAO_INFORMADO", Err, iIndice)

        Case 35930
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_COMISSAO_GRID_NAO_PREENCHIDO", Err, iIndice)

        Case 35931
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_TOTAL_COMISSAO_INVALIDO", Err, dTotalValorComissao, dValorTotal - dValorIRRF)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148825)
     
    End Select
     
    Exit Function

End Function

Function CondicaoPagto_CalculaParcelas(ByVal objCondicaoPagto As ClassCondicaoPagto, Optional ByVal bCalculaValores As Boolean = True, Optional ByVal bCalculaDatas As Boolean = True, Optional ByVal bLerDoBD As Boolean = True, Optional ByVal bTrataTeto As Boolean = False) As Long
'calcula os valores e datas de vencimentos das parcelas
'ATENCAO: a condicao de pagto tem que ter sido carregada do bd, inclusive preenchendo colParcelas,
'e preenchidas as datas de emissao, entrega, ref e o valortotal
'ATENCAO 2: as "parcelas" com % zero de recebimento sao excluidas da colecao

Dim lErro As Long, objParc As ClassCondicaoPagtoParc, dtDataParcAnterior As Date, dtDataParc As Date
Dim dValorAux As Double, dValorAcum As Double, dDif As Double, iIndice As Integer, iDiasAux As Integer
Dim bParcelasIguais As Boolean, dPercRecUnico As Double
Dim iResto As Integer, iNumParcelas As Integer
Dim dValorParcela As Double, dtDataEmissao As Date, dtDataRef As Date, dtDataEntrega
Dim iIndice1 As Integer
Dim dtDataVenctoReal As Date
Dim objParcAux As ClassCondicaoPagtoParc, colParcelasAux As New Collection
Dim dDivParc As Double, iParc As Integer, dResto As Double

On Error GoTo Erro_CondicaoPagto_CalculaParcelas
       
    If bLerDoBD Then
    
        Set objCondicaoPagto.colParcelas = New Collection
        
        lErro = CF("CondicaoPagto_Le_Parcelas", objCondicaoPagto)
        If lErro <> SUCESSO Then gError 124054
    
    End If

    If objCondicaoPagto.dTetoParcela > 0 And bTrataTeto Then bCalculaValores = True 'Se tem teto precisa calcular os valores para ter as datas

    If bCalculaValores Then
    
        bParcelasIguais = True
        dPercRecUnico = -1
        For Each objParc In objCondicaoPagto.colParcelas
        
            If objParc.dPercReceb <> 0 Then
                If dPercRecUnico = -1 Then
                    dPercRecUnico = objParc.dPercReceb
                Else
                    If Abs(dPercRecUnico - objParc.dPercReceb) > 0.0001 Then
                        bParcelasIguais = False
                        Exit For
                    End If
                End If
            End If
            
        Next
        
        'gerar valores das parcelas
        
        If bParcelasIguais Then
        
            iNumParcelas = objCondicaoPagto.iNumeroParcelas
            
            If iNumParcelas <> 1 Then
            
                'Calcula o resto da divisão inteira
                iResto = Resto(objCondicaoPagto.dValorTotal * 100, iNumParcelas)
            
                If (iResto <> 0) Then
                    dValorParcela = (objCondicaoPagto.dValorTotal * 100) / iNumParcelas
                    dValorParcela = Int(dValorParcela) / 100
                Else
                    dValorParcela = objCondicaoPagto.dValorTotal / iNumParcelas
                End If
                
            Else
            
                dValorParcela = objCondicaoPagto.dValorTotal
                iResto = 0
                
            End If
        
            'Atualizar os valores das parcelas na coleção
            iIndice = 1
            For Each objParc In objCondicaoPagto.colParcelas
            
                If objParc.dPercReceb <> 0 Then
                
                    objParc.dValor = dValorParcela
                    
                    If iIndice = (iNumParcelas - iResto) Then
                    
                        'Soma 0.01 ao Valor da Parcela
                        dValorParcela = dValorParcela + 0.01
                        
                    End If
                    
                    iIndice = iIndice + 1
                
                Else
                
                    objParc.dValor = 0
                    
                End If
                
            Next
                    
        Else
        
            For Each objParc In objCondicaoPagto.colParcelas
            
                dValorAux = Int(100 * objCondicaoPagto.dValorTotal * objParc.dPercReceb) / 100
                objParc.dValor = dValorAux
                dValorAcum = dValorAcum + dValorAux
            
            Next
            
            dDif = Round(objCondicaoPagto.dValorTotal - dValorAcum, 2)
                
            Do While Abs(dDif) >= DELTA_VALORMONETARIO
            
                For iIndice = objCondicaoPagto.colParcelas.Count To 1 Step -1
                        
                    Set objParc = objCondicaoPagto.colParcelas.Item(iIndice)
                    If objParc.dValor <> 0 Then
                        
                        objParc.dValor = objParc.dValor + 0.01
                        dDif = dDif - 0.01
                        If Abs(dDif) < DELTA_VALORMONETARIO Then Exit For
                        
                    End If
                    
                Next
            
            Loop
    
        End If
        
    End If
    
    If bCalculaDatas Then
    
        'gerar datas das parcelas
        
        dtDataRef = objCondicaoPagto.dtDataRef
        If dtDataRef = DATA_NULA Then dtDataRef = objCondicaoPagto.dtDataEmissao
        If dtDataRef = DATA_NULA Then dtDataRef = gdtDataAtual
        
        dtDataEmissao = objCondicaoPagto.dtDataEmissao
        If dtDataEmissao = DATA_NULA Then dtDataEmissao = dtDataRef
        
        dtDataEntrega = objCondicaoPagto.dtDataEntrega
        If dtDataEntrega = DATA_NULA Then dtDataEntrega = dtDataRef
        
        dtDataParcAnterior = DATA_NULA
        
        For Each objParc In objCondicaoPagto.colParcelas
        
            Select Case objParc.iTipoDataBase
            
                Case CONDPAGTO_TIPODATABASE_EMISSAO
                    dtDataParc = dtDataEmissao
                
                Case CONDPAGTO_TIPODATABASE_ENTREGA
                    dtDataParc = dtDataEntrega
                
                Case CONDPAGTO_TIPODATABASE_PARCANTERIOR
                    dtDataParc = dtDataParcAnterior
                
                Case Else
                    dtDataParc = dtDataRef
                        
            End Select
            
            Select Case objParc.iModificador
            
                Case CONDPAGTO_MODIFICADOR_FORAMES
                    dtDataParc = DateAdd("m", 1, CDate("01/" & Month(dtDataParc) & "/" & Year(dtDataParc))) - 1
                
                Case CONDPAGTO_MODIFICADOR_FORAQUINZENA
                    If Day(dtDataParc) <= 15 Then
                        dtDataParc = CDate(CDate("15/" & Month(dtDataParc) & "/" & Year(dtDataParc))) - 1
                    Else
                        dtDataParc = DateAdd("m", 1, CDate("01/" & Month(dtDataParc) & "/" & Year(dtDataParc))) - 1
                    End If
                
                Case CONDPAGTO_MODIFICADOR_FORASEMANA
                        dtDataParc = DateAdd("d", 8 - Weekday(dtDataParc), dtDataParc)
                    
            End Select
            
            Select Case objParc.iTipoIntervalo
            
                Case CONDPAGTO_TIPOINTERVALO_DIAS
                    dtDataParc = dtDataParc + objParc.iDias
                    
                Case CONDPAGTO_TIPOINTERVALO_DIAS_UTEIS
                    For iIndice1 = 1 To objParc.iDias
                        dtDataParc = dtDataParc + 1
                    
                        lErro = CF("DataVencto_Real", dtDataParc, dtDataVenctoReal)
                        If lErro <> SUCESSO Then gError 126516
                        
                        dtDataParc = dtDataVenctoReal
                    Next
                    
                Case CONDPAGTO_TIPOINTERVALO_DIAMES
                
                    iDiasAux = objParc.iDias - Day(dtDataParc)
                    If iDiasAux > 0 Then
                        
                        dtDataParc = dtDataParc + iDiasAux
                    
                    Else
                    
                        dtDataParc = dtDataParc + iDiasAux
                        dtDataParc = DateAdd("m", 1, dtDataParc)
                        'Dia 30/01 + 1 mês = 29/02 o que fazia com q no mês seguinte iDiasAux > 0 (=1) gerando outra parcela no dia 01/03,
                        'então se ao ir para o mês seguinte continuar com o dia anterior é porque o mês acaba antes então pega o 1o dia
                        'do mês seguinte para manter a sequência certa, ou seja, 30/01, 01/03 e 30/03 ao invés de 30/01, 29/02 e 01/03
                    
                        iDiasAux = objParc.iDias - Day(dtDataParc)
                        If iDiasAux > 0 Then
                            dtDataParc = dtDataParc + 1
                        End If
                        'Fim
                    
                    End If
            
                Case CONDPAGTO_TIPOINTERVALO_DIASEMANA
                    '??? completar
            
            End Select
            
            objParc.dtVencimento = dtDataParc
            dtDataParcAnterior = dtDataParc
            
        Next
    
    End If
    
    iIndice = 1
    
    'elimina as parcelas que nao tem valor
    For Each objParc In objCondicaoPagto.colParcelas
    
        If objParc.dPercReceb = 0 Then
        
            Call objCondicaoPagto.colParcelas.Remove(iIndice)
        
        Else
        
            iIndice = iIndice + 1
        
        End If
        
    Next
    
    'Ajustas as parcelas acima do teto
    If objCondicaoPagto.dTetoParcela > 0 And bTrataTeto Then
    
        For Each objParc In objCondicaoPagto.colParcelas
    
            If objParc.dValor > objCondicaoPagto.dTetoParcela Then
                dDivParc = objParc.dValor / objCondicaoPagto.dTetoParcela
                dResto = objParc.dValor
                If dDivParc - CInt(dDivParc) > 0 Then
                    iParc = CInt(dDivParc) + 1
                Else
                    iParc = CInt(dDivParc)
                End If
                For iIndice = 1 To iParc
                    'Última parcela
                    If iParc = iIndice Then
                        dValorAux = dResto
                        dResto = 0
                    Else
                        dValorAux = Arredonda_Moeda(objParc.dValor / iParc)
                        dResto = dResto - dValorAux
                    End If
                    
                    Set objParcAux = New ClassCondicaoPagtoParc
                    
                    objParcAux.dPercReceb = objParc.dPercReceb
                    objParcAux.dtVencimento = objParc.dtVencimento
                    objParcAux.dValor = dValorAux
                    objParcAux.iCodigo = objParc.iCodigo
                    objParcAux.iDias = objParc.iDias
                    objParcAux.iModificador = objParc.iModificador
                    objParcAux.iSeq = objParc.iSeq
                    objParcAux.iTipoDataBase = objParc.iTipoDataBase
                    objParcAux.iTipoIntervalo = objParc.iTipoIntervalo
                    
                    colParcelasAux.Add objParcAux
                Next
            Else
                colParcelasAux.Add objParc
            End If
    
        Next
        
        For iIndice = objCondicaoPagto.colParcelas.Count To 1 Step -1
            Call objCondicaoPagto.colParcelas.Remove(iIndice)
        Next
        
        For Each objParcAux In colParcelasAux
            objCondicaoPagto.colParcelas.Add objParcAux
        Next
        
    End If
    
    CondicaoPagto_CalculaParcelas = SUCESSO
     
    Exit Function
    
Erro_CondicaoPagto_CalculaParcelas:

    CondicaoPagto_CalculaParcelas = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148826)
     
    End Select
     
    Exit Function

End Function

Function Calcula_ValorPresente(ByVal objCondicaoPagto As ClassCondicaoPagto, ByVal dValorFuturo As Double, ByVal dTaxaFinanceira As Double, dValorPresente As Double, ByVal dtDataAVista As Date) As Long
'calcula o valor presente de acordo com o preço unitário e a cond pagto passados

Dim iIndice As Integer
Dim dSomaSerie As Double
Dim lErro As Long

On Error GoTo Erro_Calcula_ValorPresente

    'Lê a condicao de Pagamento
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO Then gError 62731
    
    objCondicaoPagto.dValorTotal = dValorFuturo
    objCondicaoPagto.dtDataRef = dtDataAVista
    objCondicaoPagto.dtDataEmissao = dtDataAVista
    objCondicaoPagto.dtDataEntrega = dtDataAVista
    
    'Calcula as datas de vencimento das parcelas p\ esse cond pagto
    lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto)
    If lErro <> SUCESSO Then gError 62732
    
    '******     INICIO DO CÁLCULO DO VALOR PRESENTE  **********
    dSomaSerie = 0

    For iIndice = 1 To objCondicaoPagto.iNumeroParcelas
        dSomaSerie = dSomaSerie + (objCondicaoPagto.colParcelas(iIndice).dValor / ((1 + dTaxaFinanceira) ^ ((objCondicaoPagto.colParcelas(iIndice).dtVencimento - dtDataAVista) / 30)))
    Next

    dValorPresente = dSomaSerie
    
    '***********  FIM DO CÁLCULO DO VALOR PRESENTE  ***********

    Calcula_ValorPresente = gErr

    Exit Function
    
Erro_Calcula_ValorPresente:

    Calcula_ValorPresente = SUCESSO

    Select Case gErr
    
        Case 62731, 62732
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 148827)
        
    End Select

    Exit Function

End Function

Function Calcula_Multa_Juros_Parcela2(ByVal dSaldo As Double, ByVal dPercMulta As Double, ByVal dTaxaJuros As Double, ByVal dtDataBaixa As Date, ByVal dtDataVencimentoReal As Date, ByVal dtDataVencimento As Date, dValorMulta As Double, dValorJuros As Double) As Long
'Obtem o valor da multa e dos juros para pagto da parcela na data informada

Dim lErro As Long
Dim iDias As Integer, dValorJurosDia As Double

On Error GoTo Erro_Calcula_Multa_Juros_Parcela2
    
    If dtDataBaixa > dtDataVencimentoReal Then
                
        iDias = dtDataBaixa - dtDataVencimento
        
        dValorMulta = dSaldo * dPercMulta
        
        'com juros simples
        dValorJurosDia = dSaldo * dTaxaJuros / 30
        dValorJuros = Round(dValorJurosDia * iDias, 2)
        
    Else
               
        dValorMulta = 0
        dValorJuros = 0
                       
    End If

    Calcula_Multa_Juros_Parcela2 = SUCESSO
    
    Exit Function

Erro_Calcula_Multa_Juros_Parcela2:

    Calcula_Multa_Juros_Parcela2 = Err
    
    Select Case Err
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148828)
    
    End Select
    
    Exit Function

End Function
