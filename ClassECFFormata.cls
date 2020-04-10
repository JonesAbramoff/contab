VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassECFFormata"
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
Dim sSituacao As String
Dim sDestinacao As String
Dim sApropriacao As String
Dim sControleEstoque As String
Dim sGerencial As String
Dim sTipoDocReserva As String
Dim sStatus As String
Dim lNumIntNF As Long
Dim dPrazoMedio As Double
Dim dQuantOP As Double
Dim lNumIntDocLote As Long
'**********************************
'Alteracao Daniel em 11/07/2002
Dim sSituacaoKit As String
'**********************************
Dim sCgc As String

On Error GoTo Erro_Browse_Formata_Campo

    Select Case objValorCampo.iTipo

        Case ADM_TIPO_DOUBLE
        
            Select Case objValorCampo.iSubTipo
        
                Case ADM_SUBTIPO_PERCENTUAL
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Percent")
        
                Case ADM_SUBTIPO_HORA
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CDate(vCampo)
        
                Case ADM_SUBTIPO_PRECOUNITARIO
                    If Len(Trim(gobjCOM.sFormatoPrecoUnitario)) > 0 Then
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, gobjCOM.sFormatoPrecoUnitario)
                    Else
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Standard")
                    End If
                    
                Case Else
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Standard")
                    
            End Select
        
        
        Case ADM_TIPO_DATE
            sData = Format(vCampo, "dd/mm/yyyy")
        
            If sData = "07/09/1822" Then sData = ""
            
            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sData
    
        Case ADM_TIPO_SMALLINT
    
            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CStr(vCampo)
                    
    
        Case Else
            
            Select Case objValorCampo.iSubTipo
        
                Case ADM_SUBTIPO_CONTA
                
                    If Len(vCampo) > 0 Then
                    
                        'se o campo for uma conta, mascara a conta antes de exibi-la
                        sContaMascarada = String(STRING_CONTA, 0)
                
                        sConta = vCampo
                    
                        lErro = Mascara_MascararConta(sConta, sContaMascarada)
                        If lErro <> SUCESSO Then gError 214930
                        
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
                        If lErro <> SUCESSO Then Error 214931
                        
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
                        If lErro <> SUCESSO Then gError 214929
                    
                    Else
                
                        sProdutoMascarado = ""
                        
                    End If
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sProdutoMascarado

                Case ADM_SUBTIPO_CGC
                
                    Select Case Len(Trim(vCampo))
                    
                        Case STRING_CPF 'CPF
                                        
                            sCgc = Format(vCampo, "000\.000\.000-00; ; ; ")
                    
                        Case STRING_CGC 'CGC
                    
                            sCgc = Format(vCampo, "00\.000\.000\/0000-00; ; ; ")
                            
                    End Select
                    
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sCgc

                Case Else
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CStr(vCampo)
                    
            End Select
            
    End Select
    
    Browse_Formata_Campo = SUCESSO
    
    Exit Function
    
Erro_Browse_Formata_Campo:

    Browse_Formata_Campo = gErr

    Select Case gErr
        
        Case 214929
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, sProduto)
        
        Case 214930
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", gErr, CStr(vCampo))

        Case 214931
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", gErr, CStr(vCampo))
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 214932)

    End Select

    Exit Function

End Function

