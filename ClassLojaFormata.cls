VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLojaFormata"
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
Dim sTipoContaCcl As String
Dim sNaturezaConta As String
Dim sStatus As String
Dim sSimNao As String
Dim sTipoCartao As String

On Error GoTo Erro_Browse_Formata_Campo

    Select Case objValorCampo.iTipo

        Case ADM_TIPO_DOUBLE
        
            Select Case objValorCampo.iSubTipo
        
                Case ADM_SUBTIPO_PERCENTUAL
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Percent")
                
                Case ADM_SUBTIPO_PERCENTUAL2
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = vCampo & "%"
                    
                Case ADM_SUBTIPO_HORA
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CDate(vCampo)
        
                Case ADM_SUBTIPO_PRECOUNITARIO
                    If Len(Trim(gobjFAT.sFormatoPrecoUnitario)) > 0 Then
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, gobjFAT.sFormatoPrecoUnitario)
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
    
            Select Case objValorCampo.iSubTipo
        
                Case ADM_SUBTIPO_TIPOCONTACCL
                    Call CF("Converte_TipoContaCcl", vCampo, sTipoContaCcl)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sTipoContaCcl
        
                Case ADM_SUBTIPO_NATUREZACONTA
                    Call CF("Converte_NaturezaConta", vCampo, sNaturezaConta)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sNaturezaConta
                
                Case ADM_SUBTIPO_SIMNAO
                    Call CF("Converte_CheckBox_SimNao", vCampo, sSimNao)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sSimNao
        
                Case ADM_SUBTIPO_CAIXA_STATUS
                    Call Converte_Caixa_Status(vCampo, sStatus)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sStatus
        
                Case ADM_SUBTIPO_TIPOCARTAO
                    Call Converte_TipoCartao(vCampo, sTipoCartao)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sTipoCartao
        
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
                        If lErro <> SUCESSO Then Error 57244
                        
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
                        If lErro <> SUCESSO Then Error 57245
                        
                    Else
                    
                        sCclMascarado = ""
                        
                    End If
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sCclMascarado
                    
                    
                Case Else
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CStr(vCampo)
                    
            End Select
            
    End Select
    
    Browse_Formata_Campo = SUCESSO
    
    Exit Function
    
Erro_Browse_Formata_Campo:

    Browse_Formata_Campo = Err

    Select Case Err
        
        Case 57244
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", Err, CStr(vCampo))

        Case 57245
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, CStr(vCampo))
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150910)

    End Select

    Exit Function

End Function

'Incluída por Luiz Nogueira em 13/04/04
Sub Converte_Caixa_Status(vCampo, sStatus)
'Converte os valores 1 ou 0 do campo Status para Aberto ou fechado
    
    'Se for Zero recebe Fechada
    If vCampo = STATUS_CAIXA_FECHADO Then
        sStatus = "Fechado"
    'Se não recebe Aberta
    Else
        sStatus = "Aberto"
    End If
    
    Exit Sub

End Sub

Sub Converte_TipoCartao(vCampo, sTipoCartao)
'Converte os valores 1 ou 0 do campo Status para Aberto ou fechado
    
    'Se for Zero recebe Manual
    If vCampo = STATUS_TIPOCARTAO_MANUAL Then
        sTipoCartao = "Manual"
    'Se for 1 recebe POS
    ElseIf vCampo = STATUS_TIPOCARTAO_POS Then
        sTipoCartao = "POS"
    'Se for 2 recebe TEF
    Else
        sTipoCartao = "TEF"
    End If
    
    Exit Sub

End Sub

