VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSRVFormata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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
Dim sCgc As String
Dim dtData As Date

On Error GoTo Erro_Browse_Formata_Campo

    Select Case objValorCampo.iTipo

        Case ADM_TIPO_DOUBLE
        
            Select Case objValorCampo.iSubTipo
        
                Case ADM_SUBTIPO_PERCENTUAL
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Percent")
        
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
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Standard")
                    
            End Select
        
        
'###########################################################
'Alterado por Wagner
        Case ADM_TIPO_DATE
           
            Select Case objValorCampo.iSubTipo
    
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
                        If lErro <> SUCESSO Then gError 183730
                        
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
                        If lErro <> SUCESSO Then gError 183731
                        
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
                        If lErro <> SUCESSO Then gError 183729
                    
                    Else
                
                        sProdutoMascarado = ""
                        
                    End If
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sProdutoMascarado
                    

                Case Else
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CStr(vCampo)
                    
            End Select
            
    End Select
    
    Browse_Formata_Campo = SUCESSO
    
    Exit Function
    
Erro_Browse_Formata_Campo:

    Browse_Formata_Campo = gErr

    Select Case gErr
        
        Case 183729
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, sProduto)
            
        Case 183730
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", gErr, CStr(vCampo))

        Case 183731
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", gErr, CStr(vCampo))
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183732)

    End Select

    Exit Function

End Function

