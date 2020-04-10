Attribute VB_Name = "GeracaoDRE"
Option Explicit

Public Function Gerar_DRE_DRP(ByVal colModelos As Collection, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, ByVal iFilialEmpresa As Integer, ByVal sDiretorio As String, ByVal iGrupoEmpresarial As Integer) As Long

Dim iIndice As Integer
Dim lErro As Long
Dim objPlanilha As ClassPlanilhaExcel
Dim sModelo As String

On Error GoTo Erro_Gerar_DRE_DRP

    TelaAcompanhaBatch2.dValorAtual = 0
    TelaAcompanhaBatch2.dValorTotal = colModelos.Count
        
    If iPeriodo = 0 Then
    
        For iIndice = 1 To colModelos.Count
        
            Set objPlanilha = New ClassPlanilhaExcel
        
            sModelo = colModelos.Item(iIndice)
            
            TelaAcompanhaBatch2.Log = TelaAcompanhaBatch2.Log & "Atualizando dados do DRE para o modelo " & sModelo & ", exercício " & CStr(iExercicio) & " e filialempresa " & CStr(iFilialEmpresa) & " ... "
            
            DoEvents
            
            lErro = CF("RelDRE_Calcula", sModelo, iExercicio, iFilialEmpresa, , iGrupoEmpresarial)
            If lErro <> SUCESSO Then gError 187120
            
            TelaAcompanhaBatch2.Log = TelaAcompanhaBatch2.Log & " OK." & vbNewLine
            
            lErro = Rotina_Atualizar_TelaBatch2
            If lErro <> SUCESSO Then gError 187121
            
            TelaAcompanhaBatch2.Log = TelaAcompanhaBatch2.Log & "Gerando o arquivo em excel na pasta " & sDiretorio & " ... "
            
            DoEvents
            
            lErro = CF("DRE_Move_Dados_Formato_Excel", objPlanilha, sModelo, RELDRE, sDiretorio)
            If lErro <> SUCESSO Then gError 187122
        
            'Exporta os dados do objPlanilha para o Excel
            lErro = CF("Excel_Gera_Planilha", objPlanilha)
            If lErro <> SUCESSO Then gError 187123

            TelaAcompanhaBatch2.Log = TelaAcompanhaBatch2.Log & " OK." & vbNewLine

            lErro = Rotina_Atualizar_TelaBatch2
            If lErro <> SUCESSO Then gError 187124

        Next
        
    Else
    
        For iIndice = 1 To colModelos.Count
        
            Set objPlanilha = New ClassPlanilhaExcel
        
            sModelo = colModelos.Item(iIndice)
            
            TelaAcompanhaBatch2.Log = TelaAcompanhaBatch2.Log & "Atualizando dados do DRP para o modelo " & sModelo & ", exercício " & CStr(iExercicio) & ", período " & CStr(iPeriodo) & " e filialempresa " & CStr(iFilialEmpresa) & " ... "
            
            DoEvents
            
            lErro = CF("RelDRP_Calcula", sModelo, iExercicio, iPeriodo, iFilialEmpresa, iGrupoEmpresarial)
            If lErro <> SUCESSO Then gError 187125
            
            TelaAcompanhaBatch2.Log = TelaAcompanhaBatch2.Log & " OK." & vbNewLine
            
            lErro = Rotina_Atualizar_TelaBatch2
            If lErro <> SUCESSO Then gError 187126
            
            TelaAcompanhaBatch2.Log = TelaAcompanhaBatch2.Log & "Gerando o arquivo em excel na pasta " & sDiretorio & " ... "
            
            DoEvents
            
            lErro = CF("DRE_Move_Dados_Formato_Excel", objPlanilha, sModelo, "DRP", sDiretorio)
            If lErro <> SUCESSO Then gError 187127
        
            'Exporta os dados do objPlanilha para o Excel
            lErro = CF("Excel_Gera_Planilha", objPlanilha)
            If lErro <> SUCESSO Then gError 187128
            
            TelaAcompanhaBatch2.Log = TelaAcompanhaBatch2.Log & " OK." & vbNewLine

            lErro = Rotina_Atualizar_TelaBatch2
            If lErro <> SUCESSO Then gError 187129
        
        Next
    
    End If
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_GERACAO_DRE_SUCESSO")
    
    Gerar_DRE_DRP = SUCESSO
    
    Exit Function
    
Erro_Gerar_DRE_DRP:

    Gerar_DRE_DRP = gErr

    Select Case gErr
    
        Case 187120 To 187129
        
        Case Else
           Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187130)

    End Select
    
    Exit Function

End Function

Public Function Rotina_Atualizar_TelaBatch2()
'Atualiza tela de acompanhamento do Batch

Dim lErro As Long
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Rotina_Atualizar_TelaBatch2

    'Atualiza tela de acompanhamento do Batch
    DoEvents

    TelaAcompanhaBatch2.dValorAtual = TelaAcompanhaBatch2.dValorAtual + 0.5
    TelaAcompanhaBatch2.TotArq.Caption = CStr(Fix(TelaAcompanhaBatch2.dValorAtual))
    TelaAcompanhaBatch2.ProgressBar1.Value = CInt((TelaAcompanhaBatch2.dValorAtual / TelaAcompanhaBatch2.dValorTotal) * 100)

    If TelaAcompanhaBatch2.iCancelaBatch = CANCELA_BATCH Then

        vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_GERAR_DRE_DRP")

        If vbMsgBox = vbYes Then gError 187131

        TelaAcompanhaBatch2.iCancelaBatch = 0

    End If
    
    Rotina_Atualizar_TelaBatch2 = SUCESSO
    
    Exit Function

Erro_Rotina_Atualizar_TelaBatch2:

    Rotina_Atualizar_TelaBatch2 = gErr

    Select Case gErr

        Case 187131

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187132)

    End Select
       
    Exit Function

End Function
