VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPCPFormata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Browse_Formata_Campo(objBrowse As AdmBrowse, colRegCampo As Collection, objValorCampo As AdmValorCampo, vCampo As Variant, iProximoCampo As Integer, iLinha As Integer, iIndice As Integer, iExercicio As Integer, iPeriodo As Integer, sNomeExterno As String) As Long

Dim sProdutoMascarado As String
Dim sProduto As String
Dim sCcl As String
Dim sCclMascarado As String
Dim lErro As Long
Dim sConta As String
Dim sData As String
Dim sContaMascarada As String
Dim sRecurso As String

On Error GoTo Erro_Browse_Formata_Campo

    Select Case objValorCampo.iTipo

        Case ADM_TIPO_DOUBLE
        
            Select Case objValorCampo.iSubtipo
            
                Case ADM_SUBTIPO_PERCENTUAL
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Percent")
                    
                Case Else
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Standard")
                    
            End Select
        
        
        Case ADM_TIPO_DATE
        
            sData = Format(vCampo, "dd/mm/yyyy")
        
            If sData = "07/09/1822" Then sData = ""
        
            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sData
    
        Case ADM_TIPO_SMALLINT
    
            Select Case objValorCampo.iSubtipo
                
                Case ADM_SUBTIPO_RECURSO
                
                    If vCampo = ITEMCT_RECURSO_MAQUINA Or vCampo = CODIGO_NAO_PREENCHIDO Then
                        sRecurso = ITEMCT_RECURSO_MAQUINA & SEPARADOR & STRING_ITEMCT_RECURSO_MAQUINA
                    
                    ElseIf vCampo = ITEMCT_RECURSO_HABILIDADE Then
                        sRecurso = ITEMCT_RECURSO_HABILIDADE & SEPARADOR & STRING_ITEMCT_RECURSO_HABILIDADE
                    
                    ElseIf vCampo = ITEMCT_RECURSO_PROCESSO Then
                        sRecurso = ITEMCT_RECURSO_PROCESSO & SEPARADOR & STRING_ITEMCT_RECURSO_PROCESSO
                    
                    Else
                        sRecurso = STRING_VAZIO
                    
                    End If
                    
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sRecurso
                                
                Case Else
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CStr(vCampo)
                    
            End Select
    
        Case Else
            
            Select Case objValorCampo.iSubtipo
        
                Case ADM_SUBTIPO_CONTA
                
                    If Len(vCampo) > 0 Then
                    
                        'se o campo for uma conta, mascara a conta antes de exibi-la
                        sContaMascarada = String(STRING_CONTA, 0)
                
                        sConta = vCampo
                    
                        lErro = Mascara_MascararConta(sConta, sContaMascarada)
                        If lErro <> SUCESSO Then gError 103140
                        
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
                        If lErro <> SUCESSO Then gError 103141
                        
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
                        If lErro <> SUCESSO Then gError 103142
                    
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
        
        Case 103142
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, sProduto)
            
        Case 103140
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", gErr, CStr(vCampo))

        Case 103141
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", gErr, CStr(vCampo))
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152223)

    End Select

    Exit Function

End Function

Public Sub Converte_Situacao_Kit(vCampo As Variant, sSituacao As String)
'Devolve em sSituacao a Descricao da Situacao passada em vCampo

    Select Case vCampo
    
        Case KIT_SITUACAO_PADRAO
            sSituacao = KIT_SITUACAO_STRING_PADRAO
        
        Case KIT_SITUACAO_ATIVO
            sSituacao = KIT_SITUACAO_STRING_ATIVO
        
        Case KIT_SITUACAO_INATIVO
            sSituacao = KIT_SITUACAO_STRING_INATIVO
            
        Case Else
            sSituacao = ""
        
    End Select

End Sub

