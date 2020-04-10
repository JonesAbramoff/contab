VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalCliForn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function NFCF_Cliente_Traz_NF_Tela(ByVal objCT As Object, objNFiscal As ClassNFiscal) As Long
    NFCF_Cliente_Traz_NF_Tela = SUCESSO
End Function

Function NFCF_Cliente_Validate(ByVal objCT As Object, ByVal lCliente As Long, ByVal iFilial As Integer) As Long
    NFCF_Cliente_Validate = SUCESSO
End Function

Public Function NFCF_Limpa_Tela(ByVal objCT As Object) As Long
    NFCF_Limpa_Tela = SUCESSO
End Function

Public Function NFCF_Cliente_Filial_Validate(ByVal objCT As Object, ByVal objFilialCliente As ClassFilialCliente) As Long
    NFCF_Cliente_Filial_Validate = SUCESSO
End Function

Function NFCF_Fornecedor_Traz_NF_Tela(ByVal objCT As Object, objNFiscal As ClassNFiscal) As Long
    NFCF_Fornecedor_Traz_NF_Tela = SUCESSO
End Function

Function NFCF_Fornecedor_Validate(ByVal objCT As Object, ByVal lCliente As Long, ByVal iFilial As Integer) As Long
    NFCF_Fornecedor_Validate = SUCESSO
End Function

Public Function NFCF_Fornecedor_Filial_Validate(ByVal objCT As Object, ByVal objFilialForn As ClassFilialFornecedor) As Long
    NFCF_Fornecedor_Filial_Validate = SUCESSO
End Function

Function NFCF_ValorDesconto_Validate(ByVal objCT As Object) As Long
    NFCF_ValorDesconto_Validate = SUCESSO
End Function

Function NFCF_Default_DataSaida(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NFCF_Default_DataSaida
    
    If gobjFAT.iExibeDataSaida = MARCADO Then
        objCT.DataSaida.PromptInclude = False
        objCT.DataSaida.Text = Format(gdtDataAtual, "dd/mm/yy")
        objCT.DataSaida.PromptInclude = True
    End If
    
    NFCF_Default_DataSaida = SUCESSO
      
    Exit Function

Erro_NFCF_Default_DataSaida:

    NFCF_Default_DataSaida = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156287)

    End Select

    Exit Function
    
End Function

Function NFCF_DataSaida(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long

On Error GoTo Erro_NFCF_DataSaida
      
    objNFiscal.dtDataSaida = gdtDataAtual
    
    NFCF_DataSaida = SUCESSO
      
    Exit Function

Erro_NFCF_DataSaida:

    NFCF_DataSaida = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156288)

    End Select

    Exit Function
    
End Function

Function NFCF_Move_DataSaida_Memoria(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long

On Error GoTo Erro_NFCF_Move_DataSaida_Memoria
  
    If Len(Trim(objCT.DataSaida.ClipText)) > 0 Then
        objNFiscal.dtDataSaida = CDate(objCT.DataSaida.Text)
        objNFiscal.iSemDataSaida = 0
    Else
        objNFiscal.dtDataSaida = objNFiscal.dtDataEmissao
        objNFiscal.iSemDataSaida = 1
    End If

'horasaida
    If objNFiscal.iSemDataSaida = 0 Then
        If Len(Trim(objCT.HoraSaida.ClipText)) > 0 Then
            objNFiscal.dtHoraSaida = CDate(objCT.HoraSaida.Text)
        Else
            objNFiscal.dtHoraSaida = Time
        End If
    Else
        objNFiscal.dtHoraSaida = 0
    End If
  
    NFCF_Move_DataSaida_Memoria = SUCESSO

    Exit Function

Erro_NFCF_Move_DataSaida_Memoria:

    NFCF_Move_DataSaida_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156289)

    End Select

    Exit Function
    
End Function

Function NFCF_DataSaida2(ByVal objCT As Object, ByVal objCTNFiscalPedido As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NFCF_DataSaida2

    If gobjFAT.iExibeDataSaida = MARCADO Then
        objCTNFiscalPedido.DataSaida.PromptInclude = False
        objCTNFiscalPedido.DataSaida.Text = Format(gdtDataAtual, "dd/mm/yy")
        objCTNFiscalPedido.DataSaida.PromptInclude = True
    End If

    NFCF_DataSaida2 = SUCESSO
    
    Exit Function

Erro_NFCF_DataSaida2:

    NFCF_DataSaida2 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156290)

    End Select

    Exit Function
    
End Function

Function NFCF_Cliente_Validate2(ByVal objCT As Object, ByVal objCliente As ClassCliente, ByVal objTipoCliente As ClassTipoCliente) As Long
'??? ATENCAO: tb usada em PV (artmill)
    'Verifica se o cliente possio um desconto
    If objCliente.dDesconto > 0 Then
        objCT.gdDesconto = objCliente.dDesconto
    ElseIf objTipoCliente.dDesconto > 0 Then
        objCT.gdDesconto = objTipoCliente.dDesconto
    Else
        objCT.gdDesconto = 0
    End If
        
    NFCF_Cliente_Validate2 = SUCESSO
    
End Function

Function NFCF_Cliente_Validate3(ByVal objCT As Object, ByVal objCliente As ClassCliente) As Long
'??? ATENCAO: tb usada em PV (artmill)
Dim lErro As Long
Dim objMensagem As New ClassMensagem

On Error GoTo Erro_NFCF_Cliente_Validate3

'    If objCliente.iCodMensagem > 0 Then
'
'        objMensagem.iCodigo = objCliente.iCodMensagem
'
'        'Lê a Mensagem "default" p/ o Cliente
'        lErro = CF("Mensagem_Le", objMensagem)
'        If lErro <> SUCESSO And lErro <> 19234 Then gError 184283
'
'        If lErro = SUCESSO Then objCT.Mensagem.Text = objMensagem.sDescricao
'
'    End If
    
    NFCF_Cliente_Validate3 = SUCESSO
    
    Exit Function
    
Erro_NFCF_Cliente_Validate3:

    NFCF_Cliente_Validate3 = gErr

    Select Case gErr

        Case 184283
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184284)

    End Select
    
    Exit Function

End Function

Function NFCF_Limpa_Desconto(ByVal objCT As Object) As Long
'??? ATENCAO: tb usada em PV (artmill)
    NFCF_Limpa_Desconto = SUCESSO
End Function

Public Function NFCF_PrecoUnitario_Calcula(ByVal objCT As Object, ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal sUM As String, ByVal sClienteNomeReduzido As String, ByVal iFilialCli As Integer, ByVal iTabelaPreco As Integer, ByVal dPercAcresFin As Double, ByVal iMoeda As Integer, ByVal objRomaneioGrade As ClassRomaneioGrade, dPrecoUnitario As Double, Optional ByVal bAvisaFaltaCotacaoMoeda As Boolean = False, Optional dPercDesconto As Double, Optional dPrecoComDesconto As Double) As Long
'calcula o preço unitario
'para nf e pv que só podem ser em R$ passar iMoeda como zero

Dim lErro As Long, lCliente As Long
Dim objCotacao1 As New ClassCotacaoMoeda
Dim objCotacaoAnterior1 As New ClassCotacaoMoeda
Dim objCotacao2 As New ClassCotacaoMoeda
Dim objCotacaoAnterior2 As New ClassCotacaoMoeda
Dim dCotacao1 As Double, dCotacao2 As Double
Dim vbMsgResult As VbMsgBoxResult, dFator As Double
Dim objCliente As New ClassCliente, objProduto As New ClassProduto
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem, objTabelaPreco As New ClassTabelaPreco

On Error GoTo Erro_NFCF_PrecoUnitario_Calcula

    dPrecoUnitario = 0
    
    If gobjCRFAT.iPrecoUnitarioFormPreco <> 0 And (iMoeda = MOEDA_REAL Or iMoeda = 0) Then
    
        'Verifica se o Cliente foi preenchido
        If Len(Trim(sClienteNomeReduzido)) > 0 Then
    
            objCliente.sNomeReduzido = sClienteNomeReduzido
    
            'Lê o Cliente através do Nome Reduzido
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError ERRO_SEM_MENSAGEM
    
            If lErro = SUCESSO Then lCliente = objCliente.lCodigo
                
        Else
        
            lCliente = 0
            
        End If
        
        lErro = CF("PrecoUnitario_CalculaFormPreco", iFilialEmpresa, lCliente, iFilialCli, sProduto, sUM, iTabelaPreco, dPrecoUnitario)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    If dPrecoUnitario = 0 And iTabelaPreco <> 0 Then
    
        If gobjCRFAT.iSeparaItensGradePrecoDif = MARCADO And Not (objRomaneioGrade Is Nothing) Then
        
            Set objTabelaPrecoItem = objRomaneioGrade.objTabelaPrecoItem
            
            If objTabelaPrecoItem.dPreco < DELTA_VALORMONETARIO Then
                lErro = 28014
            Else
                lErro = SUCESSO
            End If
            
        Else
        
            objTabelaPrecoItem.sCodProduto = sProduto
            objTabelaPrecoItem.iCodTabela = iTabelaPreco
            objTabelaPrecoItem.iFilialEmpresa = iFilialEmpresa
    
            'Lê a Tabela preço para filialEmpresa
            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then gError ERRO_SEM_MENSAGEM
    
            'Se não encontrar
            If lErro = 28014 Then
                objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
                'Lê a Tabela de Preço a nível de Empresa toda
                lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
                If lErro <> SUCESSO And lErro <> 28014 Then gError ERRO_SEM_MENSAGEM
    
            End If

        End If
        
        If lErro = SUCESSO Then
        
            dPrecoUnitario = objTabelaPrecoItem.dPreco
            dPrecoComDesconto = objTabelaPrecoItem.dPrecoComDesconto
            dPercDesconto = objTabelaPrecoItem.dPercDesconto

            objTabelaPreco.iCodigo = objTabelaPrecoItem.iCodTabela
            lErro = CF("TabelaPreco_Le", objTabelaPreco)
            If lErro <> SUCESSO And lErro <> 28004 Then gError ERRO_SEM_MENSAGEM
            If lErro = SUCESSO And objTabelaPreco.iMoeda <> iMoeda Then
            'Se possuem moedas diferentes precisa calcular/recalcular a cotação
                If objTabelaPreco.iMoeda <> MOEDA_REAL Then
                
                    objCotacao1.dtData = gdtDataAtual
                    objCotacao1.iMoeda = objTabelaPreco.iMoeda
                    objCotacaoAnterior1.iMoeda = objTabelaPreco.iMoeda
                
                    'Chama função de leitura
                    lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao1, objCotacaoAnterior1)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                Else
                    objCotacao1.dValor = 1
                End If
                
                If iMoeda <> MOEDA_REAL Then
                
                    objCotacao2.dtData = gdtDataAtual
                    objCotacao2.iMoeda = iMoeda
                    objCotacaoAnterior2.iMoeda = iMoeda
                
                    'Chama função de leitura
                    lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao2, objCotacaoAnterior2)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                Else
                    objCotacao2.dValor = 1
                End If
                
                If objCotacao1.dValor <> 0 Then
                    dCotacao1 = StrParaDbl(Format(objCotacao1.dValor, "#.0000"))
                Else
                    If bAvisaFaltaCotacaoMoeda Then
                        vbMsgResult = Rotina_Aviso(vbYesNo, "AVISO_MOEDA_SEM_COTACAO_DATA", objCotacao1.iMoeda, Format(gdtDataAtual, "dd/mm/yyyy"), Format(objCotacaoAnterior1.dValor, "#.0000"), Format(objCotacaoAnterior1.dtData, "dd/mm/yyyy"))
                        If vbMsgResult = vbNo Then gError ERRO_SEM_MENSAGEM
                    End If
                    dCotacao1 = StrParaDbl(Format(objCotacaoAnterior1.dValor, "#.0000"))
                End If
                
                If objCotacao2.dValor <> 0 Then
                    dCotacao2 = StrParaDbl(Format(objCotacao2.dValor, "#.0000"))
                Else
                    If bAvisaFaltaCotacaoMoeda Then
                        vbMsgResult = Rotina_Aviso(vbYesNo, "AVISO_MOEDA_SEM_COTACAO_DATA", objCotacao2.iMoeda, Format(gdtDataAtual, "dd/mm/yyyy"), Format(objCotacaoAnterior2.dValor, "#.0000"), Format(objCotacaoAnterior2.dtData, "dd/mm/yyyy"))
                        If vbMsgResult = vbNo Then gError ERRO_SEM_MENSAGEM
                    End If
                    dCotacao2 = StrParaDbl(Format(objCotacaoAnterior2.dValor, "#.0000"))
                End If
                
                If dCotacao1 = 0 Then gError 211631
                If dCotacao2 = 0 Then gError 211632
                
                'Se nao existe cotacao para a data informada => Mostra a última.
                dPrecoUnitario = dPrecoUnitario * dCotacao1 / dCotacao2
                dPrecoComDesconto = dPrecoComDesconto * dCotacao1 / dCotacao2
                
            End If
    
        End If
    
    End If
    
    objProduto.sCodigo = sProduto
    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 84029 '26638
    If lErro = 28030 Then gError 84030 '26639
    
    'Converte a quantidade para a UM de Venda
    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMVenda, dFator)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    dPrecoUnitario = dPrecoUnitario / dFator
    dPrecoComDesconto = dPrecoComDesconto / dFator
    
    If dPercAcresFin <> 0 Then
        dPrecoUnitario = dPrecoUnitario * (1 + dPercAcresFin)
        dPrecoComDesconto = dPrecoComDesconto * (1 + dPercAcresFin)
    End If
        
    dPrecoUnitario = StrParaDbl(Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario))
    dPrecoComDesconto = StrParaDbl(Format(dPrecoComDesconto, gobjFAT.sFormatoPrecoUnitario))
    
    NFCF_PrecoUnitario_Calcula = SUCESSO
    
    Exit Function
    
Erro_NFCF_PrecoUnitario_Calcula:

    NFCF_PrecoUnitario_Calcula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 84030
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 211631
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_SEM_COTACAO", gErr, objCotacao1.iMoeda)

        Case 211632
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_SEM_COTACAO", gErr, objCotacao2.iMoeda)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201580)

    End Select
    
    Exit Function

End Function
