VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATFormata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const Cod_Tipo_Ped = 1
Private Const Nome_Tipo_Ped = "Pedido"
Private Const Nome_Tipo_Orc = "Orçamento"

Function Browse_Formata_Campo(objBrowse As AdmBrowse, colRegCampo As Collection, objValorCampo As AdmValorCampo, vCampo As Variant, iProximoCampo As Integer, iLinha As Integer, iIndice As Integer, iExercicio As Integer, iPeriodo As Integer, sNomeExterno As String) As Long

Dim sContaMascarada As String
Dim sCclMascarado As String
Dim sConta As String
Dim sCcl As String
Dim lErro As Long
Dim sData As String
Dim sProdutoEnxuto As String
Dim sProduto As String
Dim sProdutoMascarado As String
Dim sTipoContaCcl As String
Dim sNaturezaConta As String
Dim sNaturezaProduto As String
Dim sStatus As String
Dim sCheckBoxSimNao As String
Dim sTipoOperacao As String
Dim sTipoAjudaCusto As String
Dim sTipoPedido As String
Dim sPositivoNegativo As String
Dim objItemPV As ClassItemPedido
Dim lStatus As Long
Dim objPedidoVenda As ClassPedidoDeVenda
Dim dQuantFaturada As Double
'###############################
'Inserido por Wagner
Dim sDataHebr As String
Dim sFamiliaMembro As String
Dim dtData As Date
Dim iAnoHebr As Integer
Dim iMesHebr As Integer
Dim iDiaHebr As Integer
'###############################
Dim sCgc As String
Dim sDocOrigem As String
Dim sProjeto As String
Dim sProjetoMasc As String
Dim objFilialEmpresa As New AdmFiliais
Dim iTipo As Integer
Dim sOBS As String, objCamposGenericosValores As New ClassCamposGenericosValores

On Error GoTo Erro_Browse_Formata_Campo

    Select Case objValorCampo.iTipo

        Case ADM_TIPO_DOUBLE
        
            Select Case objValorCampo.iSubTipo
        
                Case ADM_SUBTIPO_PERCENTUAL
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Percent")
        
                Case ADM_SUBTIPO_POSITIVONEGATIVO_PERCENTUAL
                    Call Converte_PositivoNegativo(vCampo, sPositivoNegativo)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(sPositivoNegativo, "Percent")
                
                Case ADM_SUBTIPO_HORA
                    If CDate(vCampo) <> DATA_NULA Then
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CDate(vCampo)
                    Else
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = ""
                    End If
                    
                Case ADM_SUBTIPO_PRECOUNITARIO
                    If Len(Trim(gobjFAT.sFormatoPrecoUnitario)) > 0 Then
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, gobjFAT.sFormatoPrecoUnitario)
                    Else
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Standard")
                    End If
                    
                Case Else
                    'objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Standard")
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "#,##0.00##")

                    
            End Select
        
        
'###########################################################
'Alterado por Wagner
        Case ADM_TIPO_DATE
           
            Select Case objValorCampo.iSubTipo
    
                Case ADM_SUBTIPO_DATAHEBR

                    dtData = vCampo
                    
                    If dtData = DATA_NULA Then
                        sDataHebr = ""
                    Else
                        sDataHebr = GregToHeb(dtData, iAnoHebr, iMesHebr, iDiaHebr)
                    End If
                    
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sDataHebr
    
                Case Else
                
                    sData = Format(vCampo, "dd/mm/yyyy")
                
                    If sData = "07/09/1822" Then sData = ""
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sData
                
            End Select
'###########################################################
    
        Case ADM_TIPO_SMALLINT
    
            Select Case objValorCampo.iSubTipo
        
                Case ADM_SUBTIPO_STATUS
                    Call CF("Converte_Status", vCampo, sStatus)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sStatus

                Case ADM_SUBTIPO_TIPOCONTACCL
                    Call CF("Converte_TipoContaCcl", vCampo, sTipoContaCcl)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sTipoContaCcl
        
                Case ADM_SUBTIPO_NATUREZACONTA
                    Call CF("Converte_NaturezaConta", vCampo, sNaturezaConta)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sNaturezaConta
        
                Case ADM_SUBTIPO_NATUREZAPRODUTO
                    Call CF("Converte_NaturezaProduto", vCampo, sNaturezaProduto)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sNaturezaProduto
        
                Case ADM_SUBTIPO_SIMNAO
                    Call CF("Converte_CheckBox_SimNao", vCampo, sCheckBoxSimNao)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sCheckBoxSimNao
                
                Case ADM_SUBTIPO_TIPOOPERACAO
                    Call Converte_TipoOperacao(vCampo, sTipoOperacao)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sTipoOperacao
        
                Case ADM_SUBTIPO_TIPOAJUDACUSTO
                    Call Converte_AjudaCusto(vCampo, sTipoAjudaCusto)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sTipoAjudaCusto
                
                Case ADM_SUBTIPO_TIPOPEDIDO
                    Call Converte_Pedido(vCampo, sTipoPedido)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sTipoPedido
                
                Case ADM_SUBTIPO_PERCENTUAL
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Percent")

                Case ADM_SUBTIPO_STATUS_PV
                   
                    Set objPedidoVenda = New ClassPedidoDeVenda
                   
                    lStatus = vCampo
                
                    If lStatus = 1 Then
                    
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = "BAIXADO"
                        
                    Else
                           
                        objPedidoVenda.lNumIntDoc = vCampo

                        'le o periodo do BD e exibe o nome externo
                        lErro = CF("ItensPedido_Le_NumIntDoc", objPedidoVenda)
                        If lErro <> SUCESSO And lErro <> 133084 Then gError 133086

                        For Each objItemPV In objPedidoVenda.colItensPedido
                            dQuantFaturada = dQuantFaturada + objItemPV.dQuantFaturada
                        Next

                        If dQuantFaturada = 0 Then
                            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = "ABERTO"
                        Else
                            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = "PARCIAL"
                        End If
                       
                    End If
                                           
                Case ADM_SUBTIPO_FILIALEMPRESA
                
                    objFilialEmpresa.iCodFilial = vCampo
                
                    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                    If lErro <> SUCESSO Then gError 195368
                    
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = objFilialEmpresa.sNome
                                           
                Case ADM_SUBTIPO_TIPOHARMONIA
                                           
                    iTipo = vCampo
                    
                    If iTipo = 1 Then
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = "E"
                    ElseIf iTipo = 2 Then
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = "S"
                    Else
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = vCampo
                    End If
                        
                Case ADM_SUBTIPO_STATUS_OV
                    lStatus = vCampo
                    If lStatus = 0 Then
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = vCampo
                    Else
                    
                        objCamposGenericosValores.lCodCampo = CAMPOSGENERICOS_STATUSOV
                        objCamposGenericosValores.lCodValor = lStatus
                        lErro = CF("CamposGenericosValores_Le_CodCampo_CodValor", objCamposGenericosValores)
                        If lErro <> SUCESSO And lErro <> 102399 Then gError 184763
                        If lErro = SUCESSO Then
                            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = objCamposGenericosValores.sValor
                        Else
                            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = vCampo
                        End If
                    End If
                    
                Case ADM_SUBTIPO_STATUS_OV_COMERCIAL
                    lStatus = vCampo
                
                    Select Case lStatus
                    
                        Case OV_STATUS_COMERCIAL_NAO_COTAR
                            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = STRING_OV_STATUS_COMERCIAL_NAO_COTAR
                                                    
                        Case OV_STATUS_COMERCIAL_EM_COTACAO
                            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = STRING_OV_STATUS_COMERCIAL_EM_COTACAO
                        
                        Case OV_STATUS_COMERCIAL_COTADO
                            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = STRING_OV_STATUS_COMERCIAL_COTADO
                            
                        Case OV_STATUS_COMERCIAL_REVISADO
                            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = STRING_OV_STATUS_COMERCIAL_REVISADO
                            
                        Case OV_STATUS_COMERCIAL_LIBERADO
                            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = STRING_OV_STATUS_COMERCIAL_LIBERADO
                            
                        Case Else
                            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = vCampo
                                                            
                    End Select
                    
                Case Else
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CStr(vCampo)
                    
            End Select
    
        Case Else
            
            Select Case objValorCampo.iSubTipo
        
                Case ADM_SUBTIPO_CONTA
                
                    If Len(vCampo) > 0 Then
                    
                        'se o campo for uma conta, mascara a conta antes de exibi-la
                        sContaMascarada = String(STRING_CONTA, 0)
                
                        sConta = vCampo
                    
                        lErro = Mascara_MascararConta(sConta, sContaMascarada)
                        If lErro <> SUCESSO Then gError 24305
                        
                    Else
                
                        sContaMascarada = ""
                        
                    End If
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sContaMascarada
        
                Case ADM_SUBTIPO_CCL
            
                    If Len(vCampo) > 0 Then
            
                        'mascara o centro de custo
                        sCclMascarado = String(STRING_CCL, 0)
                    
                        sCcl = vCampo
        
                        lErro = Mascara_MascararCcl(sCcl, sCclMascarado)
                        If lErro <> SUCESSO Then gError 24306
                        
                    Else
                    
                        sCclMascarado = ""
                        
                    End If
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sCclMascarado
                    
                    
                Case ADM_SUBTIPO_PRODUTO
                
                    If Len(vCampo) > 0 Then
                    
                        'se o campo for um produto, mascara o produto antes de exibi-lo
                
                        sProduto = vCampo
                    
                        sProdutoMascarado = String(STRING_PRODUTO, 0)

                        'coloca o Produto no formato que é exibido na tela
                        lErro = Mascara_MascararProduto(sProduto, sProdutoMascarado)
                        If lErro <> SUCESSO Then gError 20773
                    
                    Else
                
                        sProdutoMascarado = ""
                        
                    End If
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sProdutoMascarado
                    
'#########################################################################
'Inserido por Wagner

                Case ADM_SUBTIPO_SEQFAMILIA
                
                    Select Case vCampo
                    
                        Case -2
                            sFamiliaMembro = "Titular"
                        
                        Case -1
                            sFamiliaMembro = "Conjuge"
                        
                        Case Is > 0
                            sFamiliaMembro = "Filho(a)"
                        
                        Case Else
                            sFamiliaMembro = ""
                        
                    End Select

                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sFamiliaMembro

'#########################################################################

                Case ADM_SUBTIPO_CGC
                
                    Select Case Len(Trim(vCampo))
                    
                        Case STRING_CPF 'CPF
                                        
                            sCgc = Format(vCampo, "000\.000\.000-00; ; ; ")
                    
                        Case STRING_CGC 'CGC
                    
                            sCgc = Format(vCampo, "00\.000\.000\/0000-00; ; ; ")
                            
                    End Select
                    
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sCgc
                
                Case ADM_SUBTIPO_TIPO_DOC_PROJETO
                
                    lErro = CF("Projeto_Le_DocInfo", vCampo, sDocOrigem)
                    If lErro <> SUCESSO Then gError 182052
                    
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sDocOrigem
                
                Case ADM_SUBTIPO_PROJETO
                
                    If Len(vCampo) > 0 Then

                        sProjeto = vCampo

                        'coloca o Produto no formato que é exibido na tela
                        lErro = Retorno_Projeto_Tela2(sProjeto, sProjetoMasc)
                        If lErro <> SUCESSO Then gError 20773
                    
                    Else
                
                        sProjetoMasc = ""
                        
                    End If
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sProjetoMasc

                Case ADM_SUBTIPO_STATUSHARMONIA
                
                    sOBS = vCampo
                    
                    If Len(Trim(sOBS)) = 0 Then
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = "I"
                    Else
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = "C"
                    End If
                
                Case Else
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CStr(vCampo)
                    
            End Select
            
    End Select
    
    lErro = CF("Browse_Formata_Campo_Customizado", objBrowse, colRegCampo, objValorCampo, vCampo, iProximoCampo, iLinha, iIndice, iExercicio, iPeriodo, sNomeExterno)
    If lErro <> SUCESSO Then gError 192873
    
    Browse_Formata_Campo = SUCESSO
    
    Exit Function
    
Erro_Browse_Formata_Campo:

    Browse_Formata_Campo = gErr

    Select Case gErr
        
        Case 20773
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, sProduto)
            
        Case 24305
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", gErr, CStr(vCampo))

        Case 24306
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", gErr, CStr(vCampo))
            
        Case 133086, 182052, 192873, 195368, 184763
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149742)

    End Select

    Exit Function

End Function

Function Browse_Formata_Campo_Customizado(objBrowse As AdmBrowse, colRegCampo As Collection, objValorCampo As AdmValorCampo, vCampo As Variant, iProximoCampo As Integer, iLinha As Integer, iIndice As Integer, iExercicio As Integer, iPeriodo As Integer, sNomeExterno As String) As Long
    Browse_Formata_Campo_Customizado = SUCESSO
End Function

Sub Converte_AjudaCusto(vCampo, sTipoAjudaCusto)
'Converte os valores para fixo e minimo
    
    'Se for 1 recebe "Fixa"
    If vCampo = MARCADO Then
        sTipoAjudaCusto = ADM_TIPO_AJUDACUSTO_FIXA
    'Se não recebe "Mínima"
    Else
        sTipoAjudaCusto = ADM_TIPO_AJUDACUSTO_MINIMA
    End If
    
    Exit Sub

End Sub

Sub Converte_TipoOperacao(vCampo, sTipoOperacao)
'Converte os valores 0, 1 ou 2 da combo em Importação, Exportação e Mercado Interno
    
    'Se for Zero recebe Importação
    If vCampo = Cod_Tipo_Imp Then
    
        sTipoOperacao = Nome_Tipo_Imp
        
    Else
    
        'Se for um recebe Exportação
        If vCampo = Cod_Tipo_Exp Then
            sTipoOperacao = Nome_Tipo_Exp
        'Senão recebe Mercado Interno
        Else
            sTipoOperacao = Nome_Tipo_MercInt
        End If
        
    End If
    
    Exit Sub

End Sub

Sub Converte_Pedido(vCampo, sTipoPedido)
'Converte o valor 1 para pedido ou então para Orçamento
    
    'Se for 1 recebe Pedido
    If vCampo = Cod_Tipo_Ped Then
    
        sTipoPedido = Nome_Tipo_Ped
        
    Else
    
        sTipoPedido = Nome_Tipo_Orc
        
    End If
    
    Exit Sub

End Sub

Sub Converte_PositivoNegativo(vCampo, sPositivoNegativo)
'Converte o valor positivo para negativo
 
        sPositivoNegativo = vCampo * (-1)
    
End Sub

'Essa função é utilizada apenas em versões customizadas e na versão 2 não deve
'ter conteúdo
Public Function Inicializa_Tab_DadosFinanceiros_Usu(ByVal objTelas As Object) As Long

    Inicializa_Tab_DadosFinanceiros_Usu = SUCESSO
    
End Function

Sub Converte_CheckBox_SimNao(vCampo, sCheckBoxSimNao)
'Converte os valores 1 ou 0 da checkbox em Sim ou Não
    
    'Se for Zero recebe Não
    If vCampo = DESMARCADO Then
        sCheckBoxSimNao = "Não"
    'Se não recebe Sim
    Else
        sCheckBoxSimNao = "Sim"
    End If
    
    Exit Sub

End Sub

