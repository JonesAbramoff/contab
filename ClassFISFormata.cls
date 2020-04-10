VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFISFormata"
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
Dim sSecao As String
Dim sCgc As String

On Error GoTo Erro_Browse_Formata_Campo

    Select Case objValorCampo.iTipo

        Case ADM_TIPO_DOUBLE
        
            Select Case objValorCampo.iSubTipo
        
                Case ADM_SUBTIPO_PERCENTUAL
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Percent")
        
                Case ADM_SUBTIPO_HORA
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CDate(vCampo)
        
                Case Else
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Standard")
                    
            End Select
        
        
        Case ADM_TIPO_DATE
            sData = Format(vCampo, "dd/mm/yyyy")
        
            If sData = "07/09/1822" Then sData = ""
        
            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sData
    
        Case ADM_TIPO_SMALLINT
    
            Select Case objValorCampo.iSubTipo
                                    
                Case ADM_SUBTIPO_SECAO
                    
                    If Len(vCampo) > 0 Then
                        
                        If vCampo = SECAO_DEBITO Then
                            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = "Débito"
                        ElseIf vCampo = SECAO_CREDITO Then
                            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = "Crédito"
                        Else
                            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = "Apuração"
                        End If
                        
                    End If
    
                Case ADM_SUBTIPO_PRECADASTRADO
                    
                    If Len(vCampo) > 0 Then
                    
                        If vCampo = TIPOREGAPURACAO_PRE_CADASTRADO Then
                            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = "Sim"
                        Else
                            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = "Não"
                        End If
                        
                    End If
                       
                Case ADM_SUBTIPO_TIPOTERC
                    
                    If Len(vCampo) > 0 Then
                    
                        If vCampo = TIPO_TERC_CLIENTE Then
                            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = "Cliente"
                        Else
                            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = "Fornecedor"
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
                        If lErro <> SUCESSO Then gError 67925
                        
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
                        If lErro <> SUCESSO Then gError 67926
                        
                    Else
                    
                        sCclMascarado = ""
                        
                    End If
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sCclMascarado
                                        
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
        
        Case 67925
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", gErr, CStr(vCampo))

        Case 67926
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", gErr, CStr(vCampo))
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150292)

    End Select

    Exit Function

End Function
